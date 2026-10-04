# MCP 服务端（2026-10-04）

**一句话**：`automan mcp` 把 AutoMan **整个**作为一个 MCP 工具交出去。Claude Code 这类 Agent 说一句
"把这三张表合并成汇总.xlsx"，AutoMan 自己规划、按层路由、逐步校验、失败回滚，最后交回一份结构化报告。
撤不回的动作、覆盖已有数据，由 MCP 客户端的界面**直接问用户**（elicitation），调用它的那个模型替不了用户回答。

已在本机 Claude Code 2.1.289 上端到端跑通（见第五节）。

---

## 一、怎么接

**源码运行**（Claude Code）：

```bash
claude mcp add automan --scope user -e PYTHONPATH=<仓库目录> ^
    -- <仓库目录>/.venv/Scripts/python.exe -m automan mcp --ws D:/办公/工作区
```

路径建议写正斜杠：在 Git Bash 这类 shell 里，反斜杠会被当成转义吞掉，记进配置的路径里一个反斜杠都不剩（`C:Users…`），
服务端起不来（2026-10-04 实际撞上过）。加好之后 `claude mcp get automan` 应显示 Connected。

**便携版**：

```bash
claude mcp add automan --scope user -- <解压目录>/automan-cli.exe mcp --ws D:/办公/工作区
```

`--ws` 是工作区，一切写操作的边界，**由启动服务端的人定，调用方改不了**。`--layers`、`--model`、
`--l2-tools`、`--l2-native-model`、`--task-timeout`（默认 1800 秒）同理，会原样透传给每一次执行。
不写 `--ws` 时用程序目录下的 `workspace`。

## 二、交出去的工具

| 工具 | 只读 | 做什么 |
|---|---|---|
| `run_task(task)` | 否 | 执行一项任务，跑完返回报告：结论、每步由哪一层做的、校验结论、L2 用了哪种方式及原因、中途问过用户什么、工作区里新增/修改/删除了哪些文件 |
| `list_workspace()` | 是 | 工作区里有哪些文件（不含 AutoMan 的备份目录和 Office 锁文件） |
| `list_runs(limit)` | 是 | 这个工作区最近的运行 |
| `get_run(run_id)` | 是 | 一次运行的逐步记录 |

`run_task` 的参数**只有任务本身**：没有"跳过确认"，没有工作区，没有模型（测试守着这一条）。
任务没做成也正常返回（`outcome=failed`），原因写在 summary 和出错那一步里；
只有空任务、已有任务在跑、执行进程起不来这三种情况才报工具错误。

## 三、设计取舍

**只交"整个任务"，不交 `excel_write_range` 这种细粒度工具。** AutoMan 的安全不在某一个通道里，在编排层：
工作区白名单、执行前快照、失败回滚、不可逆动作必须人确认、每步校验、终局判定。把 COM 写单元格直接暴露出去，
外面的模型能写，但没有人替它备份、校验、问人。细粒度能力留在里面，由 AutoMan 自己的护栏管着。

**执行放在子进程里**，和客户端走同一条路（`automan run --confirm stdio --events`）：

1. **stdout 是协议线。** 内核里任何一句 print 落到服务端的 stdout，就是一帧坏掉的 JSON-RPC。
2. **COM 绑线程。** 服务端是异步事件循环；子进程里是 L2 一直以来的单线程跑法，评测台量的也是那条路。
3. **能停。** 客户端取消、整体超时，都是杀掉子进程。
4. **三样东西各有出处**：进度读子进程吐的机器行（新增的 `automan/events.py`，只报进度、不是事实来源）；
   确认读 `confirm.py` 那份协议；结论读 trace.db —— 和客户端看到的是同一份事实。

服务端不 import 编排层、通道、沙箱、模型调用，也不 import 失败归因那几个模块（报告里不能出现规则的判词，
M4⑤ 盲标要求），有禁引测试。

**确认只能来自人。** 子进程问"要不要放行"时，服务端发 elicitation，表单只有一个「确认执行」勾选框、默认不勾。
只有"用户提交且勾了"才放行；客户端没声明 elicitation 能力、用户拒绝、取消、提交了但没勾、10 分钟没人回答、
出错 —— 全部按拒绝处理，并且**把是哪一种拒绝写进轨迹**（为此 `confirm.py` 的答案加了一个可选的 `why`
字段，只改理由、不改结论：放行仍然只认 `ok` 恰好是 `true`）。

**一次只跑一个任务**：它们共用同一个桌面和工作区。第二个调用直接报错，说明前一个是什么、什么时候开始的。

**停得干净**：

- 客户端取消调用 / 超时：杀掉执行进程，报告里说清"被打断的那一步不会自动回滚，工作区可能停在半途"、
  快照在哪；这条运行在库里会一直记作 running（和客户端一样，服务端不写库）。
- 服务端自己被直接杀掉（MCP 客户端退出时常见）：执行进程放在一个 KILL_ON_JOB_CLOSE 的 Windows 作业对象里，
  服务端一没，它跟着没 —— 不会在没人看着的情况下继续操作桌面。作业对象设了 SILENT_BREAKAWAY_OK：
  任务里打开给用户看的记事本、WPS 不受影响；沙箱进程在内核自己的作业里，执行进程一死它就跟着死。

**进度**：规划完成、每一步路由到哪一层、上浮、每步完成都会发进度通知；两步之间每 10 秒一次心跳
（有的客户端靠进度通知续调用超时）。进度值严格递增（规范要求），有测试。

**只读工具只看得见自己那个工作区**：别的工作区的运行，`get_run` 也报"没有"。

## 四、各环节用哪个模型

（本机 `.env` 配置下；MCP 不改变 AutoMan 内部的任何模型选择）

| 环节 | 模型 | 思考 |
|---|---|---|
| **调用方**（Claude Code 里的模型） | 由客户端决定 | —— 它只决定"把什么任务交给 AutoMan"，看不到也不参与 AutoMan 内部的决策和确认 |
| 规划（把任务拆成步骤） | glm-4.5-air | 关 |
| L1 写代码 | glm-4.5-air | 关 |
| L2 一次给齐（json） | glm-4.5-air | 关 |
| L2 先看再写（多轮） | glm-5.3 | 只能开，用最浅的 low |
| L3 控件操作 | glm-4.5-air | 关 |
| L4 视觉定位 | glm-4.5v | 默认（spike 实测过的那套设置） |
| 终局判定 | glm-4.5-air | 关 |
| 确认 | **人** | —— 经 MCP 客户端的 elicitation，不经任何模型 |

## 五、验证

**自动测试**（`tests/test_mcp_server.py` 22 条 + `tests/test_confirm.py` 新增 1 条）：协议这头用 SDK 的进程内客户端真跑，
执行那头是真编排层（真 `Orchestrator` / `Guard` / `Trace` / `cli.with_events` / `cli.stdio_hook`）加假模型和假通道
（`tests/_fake_mcp_child.py`），不花钱、不联网、不碰桌面。覆盖：工具清单和提示、正常跑完（步骤、文件变化、进度单调）、
失败回滚、还没开始就崩、空任务、确认的五种回答（勾了 / 没勾 / 拒绝 / 取消 / 客户端不支持）、执行中途的覆盖确认、
2026-07-28 无状态连接上问不了人就拒绝、超时杀进程、走协议的取消杀进程、一次一个、只读工具的工作区隔离、备份目录和锁文件不列出、禁引、
进度行的编解码、确认记录里有整步放行、以及真起一个 `python -m automan mcp` 走 stdio 握手（验 stdout 干净）。
另有 `tests/test_l2_finish.py`（收尾关文档）7 条、`tests/test_final_evidence.py`（终局判定看到的工作簿摘要）6 条。
全量 931 passed / 5 skipped。修完之后 local + semi 两档 11 条真跑回归 11/11。

**本机 Claude Code 2.1.289 端到端**（真模型、真 WPS；`claude -p --mcp-config … --settings …`，
确认由一个 Elicitation 钩子代替人回答，并把收到的问题原样记下来）：

| 任务 | 走的层 | 确认 | 结果 |
|---|---|---|---|
| 合并 1月.csv、2月.csv 为汇总.xlsx 并按金额降序 | L1 | 无 | ok，files.added = 汇总.xlsx |
| 工资.xlsx 的 D 列（已有数据）用公式重算实发 = 基本工资 × 0.9 | L2（auto：json 撞上覆盖拦截 → 还原文件转多轮） | 覆盖 D2:D4，钩子同意 | ok，D 列是 `=C2*0.9` 等，值 7200/8100/10800 |
| 删除工作区里的 临时.txt | L1 | 钩子拒绝 | failed，needs_confirm，文件还在 |
| 同上 | L1 | 钩子同意 | ok，files.removed = 临时.txt，轨迹记"用户已确认（经 MCP 客户端）" |
| 第二条重跑（修完下面两处之后） | L2 | 覆盖 D2:D4，钩子同意 | ok，`needs_review: false`，判定理由"D 列所有单元格都包含公式，且计算结果正确"；`closed_docs: [工资.xlsx]`，工作区里没有锁文件 |

**用户本人在交互式 Claude Code 里试用**（真确认框，人点）：删除文件两次拒绝（一次直接拒绝、一次提交了但没勾）、
一次同意，另一次 CSV 汇总，四次结果都对。不勾直接回车提交，按设计记成"提交了但没勾选，按拒绝处理"。

握手时 Claude Code 协商的是 2025-11-25，声明了 `elicitation: {form, url}`。

**截图（2026-10-05，交互式 Claude Code）**：`docs/img/mcp_confirm.png` 是确认框，`docs/img/mcp_refused.png`
是取消之后调用方的回复（"不会自己去确认，也不会换别的方式绕过它删除"）。截图时发现 Claude Code 只摊开说明的头两三行、
其余折成"… (+5 more lines)"，而原先的说明第一行是套话、要做什么在第三行，判据和快照位置全被折掉 ——
改成三行、第一行就是要做什么（测试守着"要做什么在第一行、不超过三行"）。

**便携版**：`release/build.py` 的自检新增一步，裸写 JSON-RPC 对 `automan-cli.exe mcp` 握手、列工具 —— 通过。
SDK 那串依赖让便携版多了约 11 MB 二进制（主要是 cryptography）。便携版没有带密钥跑真任务
（执行走的就是客户端一直在用的 `automan-cli.exe run`）。

**真用之后改掉的**：

1. **确认记录漏了整步放行**（用户试用发现）：同意删除之后，备注里写着"用户已确认（经 MCP 客户端）"，
   `confirmations` 却是空的，两次拒绝也一样。整步放行落在 `payload["authorized"]`（批了）或
   `status=needs_confirm` + `error="未获授权：…"`（拒了），中途的覆盖确认落在 `payload["confirmations"]`，
   报告只读了后者。现在 `stepinfo.approval_lines` 把两种合在一起进 `confirmations`（读库时生成，旧运行也对得上）。
2. **终局判定看不到公式，误报"建议人工过目"**：给判定模型的工作区摘要原先用 `pandas.read_excel` —— 只有第一张表、
   只有值。现在每张表都列（最多 4 张），写明哪些列是公式、举一例、附上文件里缓存的值（没有缓存值就照实说没算出），
   没有公式也明说"全是写死的值"。只改判定看到的证据，不改判定规则；评测台的成功率走确定性判据，不受影响，
   受影响的只有"建议人工过目"的计数。
3. **WPS 里 AutoMan 自己打开的文档跑完不关**（锁文件占着）：终局判定之后，关掉**我们打开的、已存过盘的**文档；
   用户在任务前就开着的不碰；还有没保存的修改的不关、不存、不丢（报告里 `kept_open` 列出来）；
   任务说了"打开给我看""别关"的不关；我们自己起的 WPS 实例关空了就退出。另存进工作区的也算我们打开的。
4. 确认框的三处措辞：说明文字原先是 pydantic 模型的 docstring（写给开发者看的，带 Markdown 星号）；
   "执行前会先备份"对执行中途的覆盖确认不对（那时快照已经拍了）；判据不带标签、紧跟在"删除 临时.txt"下面，
   读起来像"已经删了"。
5. L2 方式里"第一次失败的原因"截到 120 字时不带省略号，读起来像原文就断在那儿 —— 现在截断处加"…"。

## 六、已知问题与边界

1. **协议 2026-07-28（无状态那版）下问不了人，会按拒绝处理。** 那一版里服务端不能在调用中途发请求，
   要靠返回 `InputRequiredResult`、让客户端带着答案重试来问人；AutoMan 的确认发生在执行中途，要支持得把
   一次调用拆成可续的几段。眼下 Claude Code 协商的是 2025-11-25，不受影响；有测试守着"问不了就拒绝"。
2. **只在 Claude Code 上验过。** Claude Desktop 等别的客户端是否支持 elicitation 没验；不支持的客户端里，
   任何需要确认的动作都会被拒绝（这是故意的）。`claude -p` 非交互模式下没配钩子时，确认会被自动拒绝。
3. 取消 / 超时之后，那条运行在库里一直是 running；被打断的那一步不回滚（快照在，可手动还原）。
4. "一次一个"只在同一个服务端进程里成立。客户端和 MCP 同时在同一个工作区里跑任务，彼此不知道。

## 七、复现

```bash
python -m pytest tests/test_mcp_server.py tests/test_confirm.py -q
python -m release.build --no-zip          # 自检里含 MCP 握手；打包前先备份 dist 里的 runs
```

端到端：在一个临时目录里放 `mcp.json`（上面"源码运行"那条命令的 JSON 形式，`--db` 指向临时库）、
`settings.json`（`hooks.Elicitation`，matcher 为 `automan`，命令输出
`{"hookSpecificOutput": {"hookEventName": "Elicitation", "action": "accept", "content": {"approve": true}}}`），然后

```bash
claude -p "用 automan 的 run_task 工具执行：「…」" --mcp-config mcp.json --strict-mcp-config ^
    --settings settings.json --allowedTools mcp__automan__run_task
```
