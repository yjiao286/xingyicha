# Windows 桌面版部署指南（星易查）

在 Windows 上免 Python、免命令行使用星易查。两种分发形式，功能完全一致：

| 形式 | 文件 | 适用 |
|------|------|------|
| 安装包 | `XingYiCha-Setup.exe`（Release 页）/<br>`星易查-Setup.exe`（Actions 产物） | 双击安装，桌面快捷方式，带卸载程序 |
| 便携版 | `XingYiCha-portable.zip`（Release 页）/<br>`星易查-便携版.zip`（Actions 产物） | 解压即用，U 盘拷贝、内网离线机 |

> 说明：GitHub Release 附件不支持中文文件名（会被截断），所以 Release 页面
> 上是 ASCII 名（`XingYiCha-*`），内容与 Actions 产物中的中文名完全一致。

## 一、获取安装包（云端构建，无需 Windows 机器）

1. 打开 GitHub 仓库 → **Actions** 标签页 → 左侧 **Windows 构建打包**
2. 点右侧 **Run workflow** → 分支选 `main` → 点绿色按钮
3. 等待约 15-20 分钟构建完成（绿色 ✓）
4. 点进本次运行，页面底部 **Artifacts** 区域下载 `星易查-Windows` 压缩包，
   解压得到 `星易查-Setup.exe` 与 `星易查-便携版.zip`
   （Release 页面则为 ASCII 名 `XingYiCha-*`）

> 打 tag（如 `v1.0.0`）再 push，会自动创建 GitHub Release 并把两个文件挂在
> Release 页面上，方便长期分发。
>
> 私有仓库使用 GitHub 免费额度（每月 2000 分钟，Windows 按 2 倍计费，
> 一次构建约消耗 30-40 分钟额度）。

## 二、安装与使用

**安装包版**：双击 `XingYiCha-Setup.exe`（或中文名 `星易查-Setup.exe`）→ 一路下一步（可勾选创建桌面快捷方式）→
完成页勾选"立即启动"或双击桌面图标。

**便携版**：解压 zip 到任意可写目录 → 双击 `星易查.exe`。

启动后会弹出控制台窗口并自动打开默认浏览器
（`http://127.0.0.1:5001`）。**关闭控制台窗口即停止服务**，功能使用与
Web 版完全相同。历史记录与上传文件保存在 `%LOCALAPPDATA%\星易查\`，
卸载重装不丢数据。

## 三、功能完整性说明

| 功能 | Windows 表现 |
|------|-------------|
| .docx / .pdf / .txt / .xlsx | ✅ 完整支持 |
| 扫描件 PDF OCR | ✅ 完整支持（模型已打包进 exe） |
| .doc 元数据比对 | ✅ 完整支持 |
| .doc 正文提取 | ⚠️ 需另装 [LibreOffice](https://www.libreoffice.org/)（免费）。未安装时自动跳过正文，其余维度不受影响；装好后重启星易查即可生效 |

如需让局域网其他电脑访问：设置环境变量 `HOST=0.0.0.0` 后启动
（首次会弹出 Windows 防火墙放行提示）。

## 四、常见问题

- **杀毒软件报毒/拦截**：PyInstaller 打包的 exe 无数字签名，部分杀软会误报。
  在杀软中将 `星易查.exe`（或安装目录）加入信任/白名单即可。代码开源可审计。
- **双击后提示端口占用**：程序会自动改用相邻端口（5002-5010），以控制台
  窗口打印的实际地址为准。
- **想换端口**：`星易查.exe 5005`（带端口参数启动）。
- **上传文件在哪**：`%LOCALAPPDATA%\星易查\uploads`（分析完成后可清理）。
- **历史记录在哪**：`%LOCALAPPDATA%\星易查\history`。

## 五、备选：本地 Windows 机器构建（云端不可用时）

在一台装了 Python 3.11 的 Windows 机器上，双击 `packaging\build_windows.bat`
即可完成与云端完全相同的构建（产物同样在 `dist\` 下）。生成安装包需另装
[Inno Setup 6](https://jrsoftware.org/isinfo.php)，只做便携版可跳过。

## 六、Docker 部署（服务器场景）

仍按仓库根目录 `docker-compose.yml` 使用，与桌面版互不影响。
