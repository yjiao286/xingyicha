# 星易查 · 围串标风险识别分析系统（XingYiCha）

基于 Flask 的 Web 应用：上传多份投标文件，通过多维度交叉分析识别围标串标行为，
依据《中华人民共和国招标投标法实施条例》第四十条输出风险评分与 DOCX 综合报告。

## 功能特性

- **5 大分析维度**：元数据比对（创建者/修改者/硬件ID）、文本相似度（k-gram
  公共段落 + 近似重复段落）、人员比对（姓名/电话/身份证/邮箱 6 层交叉匹配）、
  报价分析（大写金额解析/分项聚类/等差数列检测）、文档结构
- **文件格式**：`.docx` / `.doc` / `.pdf`（含扫描件 OCR）/ `.txt` / `.xlsx`
- **报告输出**：风险评分 + 条款判定汇总 + 证据明细 + 两级处置建议 + 法条附录
- **历史与统计**：分析记录留存、结论分布、评分趋势等统计页

## 下载与安装

桌面版免 Python、免命令行。到
[Releases 页面](https://github.com/yjiao286/xingyicha/releases)
（当前最新 [v1.5.0](https://github.com/yjiao286/xingyicha/releases/tag/v1.5.0)）
按操作系统下载对应附件：

| 操作系统 / 环境 | 下载文件 | 安装方式 | 详细指南 |
|-----------------|----------|----------|----------|
| Windows 10/11 | `XingYiCha-Setup.exe` | 双击安装，带卸载程序与桌面快捷方式 | [Windows部署.md](docs/Windows部署.md) |
| Windows（U盘/内网离线） | `XingYiCha-portable.zip` | 解压即用，双击 `星易查.exe` | 同上 |
| macOS（M1–M4 Apple Silicon） | `XingYiCha-macOS-arm64.dmg` | 拖入 Applications；首次打开右键→打开放行 | [macOS部署.md](docs/macOS部署.md) |
| macOS（Intel） | `XingYiCha-macOS-x86_64.dmg` | 同上 | 同上 |
| Linux x86_64（Ubuntu/Debian 等主流发行版） | `XingYiCha-x86_64.AppImage` 或 `XingYiCha-linux-x86_64.tar.gz` | `chmod +x` 后运行 / 解压即用 | [Linux部署.md](docs/Linux部署.md) |
| **银河麒麟 V10（x86_64：Intel/AMD/海光/兆芯）** | `XingYiCha-linux-x86_64.tar.gz`（推荐） | 见麒麟指南（glibc/KYLSEC/离线安装注意事项） | [银河麒麟部署.md](docs/银河麒麟部署.md) |
| **银河麒麟 V10（aarch64：飞腾/鲲鹏/麒麟990）** | 暂无预编译产物 → 源码部署 | 麒麟指南第六节 | [银河麒麟部署.md](docs/银河麒麟部署.md) |
| 服务器（任意系统） | 源码 + Docker | `docker compose up -d --build` | 仓库根目录 `docker-compose.yml` |

**通用说明**：桌面版启动后自动打开浏览器访问 `http://127.0.0.1:5001`，关闭
终端/控制台窗口即停止服务；端口占用时自动改用 5002-5010。`.docx` / `.pdf` /
`.txt` / `.xlsx` / 扫描件 OCR 均开箱即用；仅老格式 `.doc` 的**正文提取**需
系统另装 LibreOffice（免费）：Windows/macOS [官网下载](https://www.libreoffice.org/)、
macOS 亦可 `brew install --cask libreoffice`、Linux/麒麟 `sudo apt install
libreoffice`。未安装时该维度自动跳过（`.doc` 元数据比对仍可用），其余不受影响。

## 源码运行（开发 / 无预编译产物平台）

```bash
python3 -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
python3 app.py 5001
```

依赖自检：`python app.py --check`。

## 开源许可

本项目采用自定义开源许可协议，全文见 [LICENSE](LICENSE)，要点：

- 个人学习、科研教学、公益活动及单位内部使用：**免费**，可自由使用、修改、
  再分发（须保留版权声明与协议全文）；
- **商业使用（销售、付费集成、有偿检测服务等）必须事先向
  中国星网易联供应链有限公司报备**，经确认后方可实施，具体授权模式以公司
  回复为准。报备邮箱：`jiaoy1@chinasatnet.com.cn`，其他渠道见
  [LICENSE](LICENSE) 第十条。

本项目含的第三方开源组件（Flask、PyMuPDF、python-docx、RapidOCR 等）遵循其
各自的开源协议。
