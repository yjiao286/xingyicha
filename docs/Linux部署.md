# Linux 桌面版部署指南（星易查）

两种分发形式，功能与 Windows 桌面版完全一致：

| 形式 | 文件 | 适用 |
|------|------|------|
| AppImage | `XingYiCha-x86_64.AppImage` | 单文件，`chmod +x` 后双击即运行 |
| 便携版 | `XingYiCha-linux-x86_64.tar.gz` | 解压即用 |

## 一、获取（云端构建，与 Windows 版同源）

1. 登录 GitHub（私有仓库需有权限的账号）
2. Actions 页手动 Run workflow，或打 `v*` tag 自动发布
3. Release 页面下载 `XingYiCha-x86_64.AppImage` 或 tar.gz
   （Actions 产物与 Release 附件内容一致）

## 二、使用

**AppImage**：
```bash
chmod +x XingYiCha-x86_64.AppImage
./XingYiCha-x86_64.AppImage        # 双击亦可（需桌面环境文件管理器设置"可执行"）
```
启动后自动打开浏览器访问 `http://127.0.0.1:5001`，终端窗口关闭即停止服务。

**便携版**：
```bash
tar xzf XingYiCha-linux-x86_64.tar.gz
./XingYiCha/XingYiCha
```

## 三、功能说明

| 功能 | 表现 |
|------|------|
| .docx / .pdf / .txt / .xlsx / 扫描件 OCR | ✅ 完整支持 |
| .doc 元数据比对 | ✅ 完整支持 |
| .doc 正文 | ⚠️ 需系统安装 LibreOffice（`apt install libreoffice` 等），未装时自动跳过，其余不受影响 |

数据目录：`~/.local/share/星易查/`（history 与 uploads）。

## 四、常见问题

- **报 libGL.so.1 缺失**：`sudo apt install libgl1`（AppImage 依赖宿主系统库）。
- **AppImage 无法双击启动**：多数文件管理器需勾选"允许作为程序执行"；命令行 `./` 运行不受限。
- **端口占用**：自动改用 5002-5010，以终端打印的实际地址为准。
- **Intel 与 ARM 平台**：当前产物为 x86_64；ARM 桌面（如树莓派）暂未构建。
