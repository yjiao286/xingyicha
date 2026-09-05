# AGENTS.md

This file provides guidance to AI coding agents (Claude Code, ZCode, Codex, Cursor, etc.) when working with code in this repository.

## 项目概述

围串标风险识别分析系统（"星易查"）— 基于 Flask 的 Web 应用，通过多维度分析投标文件来检测围标串标行为。依据《中华人民共和国招标投标法实施条例》第四十条。

## 开发命令

```bash
# 激活虚拟环境并运行
./venv/bin/python3 app.py 5001          # 开发模式 (debug=DEBUG env, default False)

# Docker 部署（生产模式：gunicorn）
docker compose up -d                    # 构建并启动，绑定 5001 端口

# 语法检查
./venv/bin/python3 -c "import py_compile; py_compile.compile('app.py', doraise=True)"

# 依赖自检（便携包目标机校验）
./venv/bin/python3 app.py --check

# 健康检查
curl -s -o /dev/null -w "%{http_code}" http://localhost:5001/

# 停止开发服务器
lsof -ti:5001 | xargs kill -9
```

环境变量：`DEBUG=1`（开启 Flask debug）、`SECRET_KEY`（随机生成）、`MAX_CONTENT_LENGTH_MB`（可选，设置后作为单次上传总上限，413 返回 JSON 错误；**默认不设上限**）、`LOG_LEVEL=INFO`、`OCR_TIME_BUDGET`/`OCR_MAX_PAGES`（扫描件 OCR 预算，**默认 0=完全放开**，可按需限制）、`ANALYSIS_TIMEOUT=3600`（整体分析超时，需 < gunicorn --timeout）。OCR 依赖（pymupdf + rapidocr_onnxruntime）缺失时扫描件自动跳过，其余功能不受影响。

## 桌面版（Windows / Linux / macOS）

`packaging/` 为 PyInstaller + Inno Setup/AppImage/dmg 构建资产，`requirements-desktop.txt`（waitress 替代 gunicorn，gunicorn 不支持 Windows）。云端构建：`.github/workflows/desktop-build.yml` 四组合 matrix（Windows / Ubuntu x86_64 / macOS arm64 / macOS x86_64 后者经 Rosetta），**push `v*` tag 自动三平台构建并发布 GitHub Release**（产物 ASCII 名 `XingYiCha-*`；Release 附件不支持中文文件名）。产物由各平台 job 用 `gh release upload` **直传 Release**（私有库 Actions Artifact 存储配额按月累计、触顶后清空也不解禁，Release 附件通道不受限）；手动 dispatch 可传 `release_tag` 参数把产物补传到指定 Release，无参数的手动运行才走 Artifact（保留 1 天）。`app.py` 冻结适配：`IS_FROZEN` 时资源目录取 `sys._MEIPASS`、数据目录按平台落 `%LOCALAPPDATA%\星易查` / `~/.local/share/星易查` / `~/Library/Application Support/星易查`、强制 stdout UTF-8、waitress + 自动开浏览器 + 端口占用回退 + 单实例探测（`/api/ping` 标记 + `_probe_own_instance`：已有实例在运行则直接拉起页面退出，防止双实例同写 history）；macOS 额外跑 Cocoa 图形层（`_run_macos_gui`：NSApplication 占主线程提供 Dock reopen 回调——点图标/Finder 重开即恢复页面，菜单栏"打开页面/退出"，waitress 转后台 daemon 线程；AppKit 缺失时回落纯控制台模式）。详见 `docs/Windows部署.md` / `docs/Linux部署.md` / `docs/macOS部署.md`。

## 部署源码包

`星易查_围串标分析系统.tar.gz`（gitignore，本地产物）为 Docker 部署用源码包，含 app.py / requirements.txt / Dockerfile / docker-compose.yml / .dockerignore / AGENTS.md / templates / static。**代码变更后需重打**：`tar -czf 星易查_围串标分析系统.tar.gz app.py requirements.txt Dockerfile docker-compose.yml .dockerignore AGENTS.md templates static`。

> **注意**：新增前端/报告使用的提取字段时，须同步加入 `_prepare_history_data()` 的 `_HISTORY_KEEP` keep 列表（app.py），否则该字段在历史记录及"历史记录 → 下载报告"场景中会丢失。keep 列表是"白名单"——除 `_` 开头键外其余键按类型置空（str→`''`、其他→`None`）；三个分区的 keep 均须包含 `name`（曾有重构丢 `name` 导致历史记录文件名全空的回归）。

## 架构

### 单体应用 (`app.py`，~6300 行)

所有后端逻辑集中于 `app.py`，无 Blueprint 或模块拆分。代码按功能区段组织，用 `# ── Section ──` 注释分隔：

| 区段 | 功能 |
|------|------|
| Helpers | `sanitize_text`、`_find_tool`、`_safe_save`、`_is_within_upload_folder` |
| Config | `CONFIG` 字典集中管理所有分析阈值 |
| .doc Conversion | LibreOffice 转换（缓存避免重复启动）；antiword/catdoc 回退 |
| Metadata Extraction | .docx（XML）、.pdf（PDF info）、.doc（OLE2 + KSO，含值格式验证） |
| Text Extraction | 纯文本 + 表格；PDF 进度回调 + 早停（连续空页）+ 扫描件 OCR 回退（pymupdf 渲染 + RapidOCR，默认放开，可用 `OCR_TIME_BUDGET`/`OCR_MAX_PAGES` 限制）+ PDF 结构化表格通道（PyMuPDF `find_tables` → 管道表行，仅关键词触发页，复用 docx/xlsx 管道表解析器，依赖缺失自动降级）；.txt 编码自动识别（UTF-8/GB18030）；.xlsx 工作表（openpyxl） |
| Personnel Extraction | 中文姓名（2-4字 + 姓氏字典校验，少数民族间隔名跳过姓氏检查）+ 英文姓名（字母/连字符/空格，2-40字）；联系方式池（电话/身份证/邮箱/银行账号）跨文档频次降权（≥3 份且 ≥80% 文档出现的值视为招标方环境噪声剔除）+ 银行账号招标方收款上下文排除；电话定位遍历授权/人员/签名/投标函（cover）区段 |
| Extraction Robustness | PDF/OCR 文本污染防御层（在 `extract_personnel`/`extract_prices` 入口统一预处理）：`_glue_phrases()`（被换行/空格从中间拆散的 ~45 个敏感词组重粘，如 `法\n定\n代\n表\n人`、`委托代\n理人`，先判文本是否存在避免空转）+ `_normalize_cjk_whitespace()`（CJK 标点周围空白收紧：`投标人 ： 张三`→`投标人： 张三`，Unicode 空格统一 + 2+ 空格折叠）+ `_join_split_names()`（姓名换行拼接，两次收敛；**仅拼换行不拼空格**，避免列间距把 `王强 联系` 拼成标签噪声）+ `_clean_company()`（残缺标签 `（投标人名称`/`（盖单位章`、下划线填充、前缀 `投标人：`/`单位名称：` 剥离）+ `_clean_phone()`（前导 `_`/尾部标点/号内列距空格清洗，`010 - 1234 5678`→`010-12345678`）；区段识别到但提取全空时全文兜底重扫（company/legal/auth 三者全空门槛，防覆盖与重复）。第二扩充批次：`_strip_invisible()`（零宽字符/BOM/软连字符删除，`\r`/`\f`/`\v`/`\u2028` 归一为换行——`\s` 不匹配零宽族，单个 `\u200b` 即可击穿全部标签正则）+ `_fix_ocr_label_confusions()`（扫描件标签形近字修复：`法定代表入`/`电活`/`人民巾`/`联糸`/`委托入` + 异体拼写 `身分证`，全部为"合法中文中不存在的字符串"，全局替换零误伤）+ `_iter_mobiles()`（手机号形态池：裸号/`+86`/`86` 前缀——原 `(?<!\d)` 会误杀 86 前缀/短横线分组 `139-1234-5678`；座机标签通道同步扩展 `移动电话/手机号码/电话号码` 并兜底无区段文档）+ 金额千分位空格分组（`_AMT_ARABIC` 容忍 `1 261 819.76` 细空格分组；`_parse_amount` 仅在"空格后恰 3 位数字"时合并，`合计 1838529 5002800` 双列号永不粘连）+ 姓名捕获容错类（`[一-鿿](?:[ \t]*[一-鿿]){1,3}` 容忍字间空格 `张 三`；无锚点捕获补 `(?![一-鿿])` 回溯护栏防吞后续 CJK 词；人员表 3 处"空格=列分隔符"模式与模式 8 严格段**保持连续类不动**（严格先行+容错重试两段式）；`_is_person_name` 中文路径先折叠空格再验证；all_persons 去重前统一折叠，跨文件同人可比；少数民族间隔点 `·` 两侧空格容忍）+ 地址截断（`地址：… 电话：…` 同行时在下一个字段标签处切断）+ 全角 `＠` 归一 |
| Price Extraction | 含全中文大写数字解析（`_cn_to_number`）；6 通道提取 + 后校验 + 行内非报价上下文排除（保证金/合同金额/最高限价/控制价等，通道遍历所有匹配防遮蔽）+ 大写↔小写互验（不一致采用小写并记 warning）+ 分项汇总验证 + `warnings[]` 溯源（前端报价页展示） |
| Text Similarity | k-gram 哈希索引 + 双向最大延伸（O(n+m)，替代全文 SequenceMatcher）+ 空白归一化；三层过滤（参照/规则/全局共现）+ 实质性评分；近似段落检测（difflib 段落对，捕获轻度编辑串标） |
| Document Structure | 扩展标题正则（中文数字/章节/Section/附录/字母编号） |
| Comprehensive Analysis | `run_full_analysis()` + 加权评分 + 协同加分 + 三档结论 |
| Report Generation | `generate_report_docx()`（概览+文件清单+维度统计+风险摘要 → 四维明细（含严重度排序/人员名单/报价汇总表/分项比对）→ 条款判定汇总表+评分构成+证据明细+评分规则 → 结论+两级处置建议（按结论等级的总体建议 + 按命中条款/银行账号线索动态生成的专项建议）+使用说明 → 附录法条；全部防御式取值，兼容历史精简数据） |
| Routes | Flask API 端点（流式 NDJSON 进度 + uuid 前缀文件名 + 上传无上限 + 路径校验 + `/api/cancel` 取消机制：`_CANCEL_EVENTS` 注册表 + `AnalysisCancelled` 异常 + 各层检查点） |

### 分析维度（5 个维度）

1. **元数据比对** — 交叉比对创建者/修改者/应用/模板/WPS 硬件ID/ICV；多卷文件聚合全部卷的元数据
2. **文本相似度** — `find_common_segments()`（k-gram 索引 + 双向最大延伸，空白归一化，支持中英混排）；三层过滤 + 四维实质性评分；`_find_near_duplicate_paragraphs()` 检测轻度编辑的近似段落；`_strip_toc_dots()` 匹配前剥离目录点线（`[.．…]{2,}`），防止无关目录行跨文档误配
3. **人员比对** — 中文 + 英文姓名提取；提取层含管道表解析（`_parse_personnel_pipe_table`，docx/xlsx 表格按列名映射姓名/职务/电话/身份证）与无章节回落；职称词黑名单（中级/教授/工程师等，PDF"姓名 职称 分工"列式表格）＋姓名-职称-角色容错（`张伟 中级 项目负责人`→张伟）＋职务标签列跳过（压平表格行 `陈刚 任中职务 项目经理` 中姓名与角色之间夹标签列时跳过取真名，容忍冒号/零间距/繁体形态 `任中职务：项目经理`、`任職務`；`_is_person_name` 通用拒绝以 职务/職務/岗位/崗位/职称/職稱/角色/职责 结尾的字符串——`任中职务` 恰以姓氏`任`开头曾混入 all_persons 并触发第三项误判）＋集体词黑名单（含 项目/人员/团队/成员/机构/分工/部门/简历/配备/配置/一览 简繁任一的字符串拒绝——章节头 `项目团队`/`关键人员`/`管理人员` 以真实姓氏 项/关/管 开头可穿透姓氏校验）＋角色配对防错位（角色-姓名配对限定同行，防上一行行尾角色抓下一行行首姓名 `王强 项目经理\n李勇 施工员`→李勇误成项目经理；姓名-角色模式后置过滤，已被前序角色词持有的姓名不与后续角色重配 `项目经理 王强 技术负责人 李四`；姓名↔职务标签跨度可跨换行但禁越下一个 `姓名：` 标签，防跨行借角色）；all_persons 内完全相同 (name, role) 对去重；授权书模式涵盖 `兹委托/现委托/现授权/特授权` + `代理人` 容忍换行（`现委托李四为我方代理\n人`）；`phones[]`/`id_numbers[]`/`emails[]` 多值池 + 6 层交叉匹配（同名/共享手机/共享身份证/邮箱/授权交叉/修改人/重叠率）；文本在进入区段检测前先过 Extraction Robustness 防御层（详见上表）
4. **报价分析** — 总价/分项/成本明细；中文大写金额解析（含独立大写通道、`￥12.5万元` 单位保留、全角数字归一、保证金上下文过滤、`bidRate` 费率/下浮率字段）；分项模糊聚类（LCS + Jaccard 2-gram）+ 等差数列检测；误提取防御层：银行凭证排除（`_amount_in_payment_voucher` 凭证词窗口检查——保证金电汇凭证 `金额/人民币：22,000.00/贰万贰仟元整/账号/开户行/制单` 的银行词在相邻行，逐行前缀检查拦不住）+ 保证金金额回声集（`投标保证金 2.2万元` 标注一次、凭证里无标注复现一次，等值候选全部跳过）+ `人民币` 前缀同行限定（`2000万元人民币\n2014年1月20日` 跨行抓年份，被验证层清空后低优先级通道不再重试）+ 数字后紧跟「年」的年份守卫 + 报价章节锚点校验（目录条目无点线前缀时靠「锚点后 400 字符内须有价格信号」拒绝；`《…》` 与 `“…”` 引用提及跳过）+ 成本行 Pattern C 万倍率（`350万`→350万而非丢单位）与业绩表上下文排除（`合同金额` 表头下的历史合同额如 `站产品采购 3500万` 不是本标成本项）；第二扩充批次：单价率守卫（`_amount_is_unit_rate`——金额后紧跟 `/月`、`/人月`、`元/次` 等分母后缀的是单价率不是总价，服务类标书按人月计价常见）＋费用/违约金标签排除（`_NON_BID_AMOUNT_CTX` 增 违约金/赔偿金/罚金/代理服务费/中标服务费/交易服务费/平台使用费/工本费/标书款/手续费/佣金）＋保证金回声集支持中文大写（`投标保证金（大写）：贰万元整` 也入回声集，统一走 `_parse_amount`）＋日期串守卫（`_is_yyyymmdd`——税率分解裸数字对模式中 `签订 20250826/生效 20250829/3` 比值恰为 1.0 可穿透比值检查；成本行粘连日期 `检测20250826批次` 同防）＋成本行形态容错（序号前缀 `1. `/`（二）`、名称与数值间冒号、千分位逗号 `340,000.00`）＋凭证词补充（承兑/汇票/支票/本票/网银）
5. **文档结构** — 扩展标题格式（第X章/Section X/附录X/字母编号）

### 前端

- `templates/index.html` — 单页面应用（吸顶导航切换「分析工作台 / 数据统计」两视图，支持 `#stats` 深链；拖拽上传、结果展示标签页）
- `static/css/style.css` — 所有样式（含统计页：KPI 卡片、SVG 环图/趋势图/维度条形图、最近记录表）
- `static/js/main.js` — 所有前端逻辑（文件上传、流式 NDJSON 进度——含 `pdf_ocr`/`pdf_ocr_start` 扫描件 OCR 进度阶段、结果渲染——含费率报价（`bidRate`）与多值联系人就展示、报告下载、`renderStats()` 统计视图）。人员交叉页：KPI stat-cards（严重度 chips 点击过滤异常卡）+ 各标书属性合并对比表（按原始值判定共享高亮、身份证/银行账号打码点击显示）、人员交叉矩阵（类型过滤 chips、悬停共享行列高亮）+ 按严重度分组可筛选的异常卡片；报价页：KPI 概览卡（文件数/最高/最低/价差%）+ 横向条形图（相同报价标红、缺失键文件以 — 展示且不参与比对）+ 明细表（含费率/下浮率行、数字右对齐等宽、万元换算副行、同值高亮）+ 分项卡可折叠（风险徽标卡默认展开；完全一致/等差数列/高度接近）
- 图表全部为手写 SVG（无 CDN 依赖，适配内网离线部署）；统计数据来自 `GET /api/stats`

### 文件格式支持

- `.docx` — 原生支持（python-docx）
- `.pdf` — 原生支持（pypdf）；无文字页面（扫描件）自动回退 OCR：pymupdf 渲染 + RapidOCR（onnxruntime），默认完全放开（`OCR_TIME_BUDGET`/`OCR_MAX_PAGES` 设为 0=不限），整体受 `ANALYSIS_TIMEOUT` 约束
- `.doc` — 文本需安装 LibreOffice（优先项）或 antiword（通过 `extract_doc_text_raw()`）；元数据通过 `olefile`（Python 依赖）读 OLE2 属性流（同 .docx 维度）
- `.xlsx` — 报价附件支持（openpyxl，`read_only`+`data_only`）；工作表输出为 `【工作表】名` 表头 + ` | ` 分隔行，复用 docx 管道表解析器；元数据走同一 OOXML zip 路径（docProps/core.xml）
- `.txt` - 纯文本支持（OCR 标书输出）；`_read_text_file()` 自动识别 UTF-8(含BOM)/GB18030 编码；无文档元数据，元数据维度按"无法判断"处理

### API 端点

| 路由 | 方法 | 说明 |
|------|------|------|
| `/` | GET | 渲染主页面 |
| `/api/upload` | POST | 上传多份标书文件（uuid 前缀防覆盖），返回文件名列表。支持 `.docx/.doc/.pdf/.txt/.xlsx` |
| `/api/analyze` | POST | 接收 `{files: [filename, ...]}`，返回完整分析结果（JSON）。路径穿越保护：仅接受纯文件名 |
| `/api/analyze_stream` | POST | 上传+分析，返回 NDJSON 流式进度事件（首个 `type: ready` 携带 `request_id`，最终 `type: result`），**前端首选**；支持取消 |
| `/api/cancel` | POST | 接收 `{id: request_id}`，置位取消 Event；提取/OCR/分析线程在检查点抛 `AnalysisCancelled`，流以 `type: cancelled` 事件正常收尾（非错误） |
| `/api/single_upload_and_analyze` | POST | 合并上传+分析（JSON 响应） |
| `/api/report` | POST | 接收 `{analysis: {...}}`，校验结构完整性，返回 .docx 报告 |
| `/api/history` | GET | 列出所有历史分析记录 |
| `/api/history/<id>` | GET | 加载指定历史记录（ID 格式白名单校验） |
| `/api/history/<id>` | DELETE | 删除指定历史记录（格式校验） |
| `/api/stats` | GET | 聚合所有历史记录生成统计页数据：结论分布、平均评分、按日趋势、四维度命中数、最近记录 |

### 关键配置

- 端口：`5001`
- 上传目录：环境变量 `UPLOAD_FOLDER` 或临时目录；文件名 uuid 前缀防并发覆盖
- 上传上限：**默认不限制**（`MAX_CONTENT_LENGTH_MB` 设置后生效；413 由 `_too_large` errorhandler 返回 JSON 便于前端提示）
- `ANALYSIS_TIMEOUT`：3600 秒（默认，< gunicorn timeout 3900；OCR 全放开后大扫描件可达数十分钟）
- `MAX_PDF_PAGES`：0 = 不限页数（默认）；`MAX_FILE_SIZE_MB=300`/`MAX_TOTAL_SIZE_MB=500` 仅 UI 警告不阻止
- `requirements.txt`：flask, python-docx, pypdf, gunicorn, olefile, openpyxl（xlsx）, pymupdf + rapidocr-onnxruntime（扫描件 OCR 回退）
- Docker 镜像基于 `python:3.11-slim`，另安装 `antiword` + `libgl1`（opencv 运行时依赖）+ `fonts-noto-cjk`（OCR 中文字形）
- 离线验证：`tools/replay_extraction.py <语料目录>` 批量回放文本/人员/报价提取（报告固定写到当前目录 `replay_report.json`）；`tests/test_extraction.py` 单元样本回归（23 例）
- 分析结果字段：`pricing.files[].bidRate`（费率/下浮率报价）；`personnel.files[].phones[]/id_numbers[]/emails[]`（多值联系池，前端与 .docx 报告均已展示）；顶层 `project_name`（跨文档投票提取的项目名——标签正则含"项目名称/工程名称/标段名称"等，值经引号书名号剥离/填空下划线剔除/标签词与日期值过滤，≥2 份文档一致优先；`_prepare_history_data` 保留顶层键，旧历史记录缺失时报告命名自动退化）。报告下载名：`围串标风险识别分析报告_[项目名_]判定等级_YYYYMMDD_HHMMSS.docx`（前端优先取 Content-Disposition 的服务端文件名）

### 参照文件功能

当用户上传招标文件/技术要求作为"参照文件"时，`text_similarity_analysis()` 会先索引参照文件中的文本段落。标书之间发现的相似段落中，在参照文件中出现过的部分会被标记为 `.is_reference_from_template = True`，从而降低其风险权重。这一过滤逻辑在 `classify_abnormal_reason()` 中生效：被过滤的匹配项会被排除或降级。
