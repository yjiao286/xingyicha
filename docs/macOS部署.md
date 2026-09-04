# macOS 桌面版部署指南（星易查）

功能与 Windows / Linux 桌面版完全一致。Apple Silicon (M 系列) 与 Intel 双架构：

| 架构 | dmg 安装包 | zip 便携版 |
|------|-----------|-----------|
| Apple Silicon (M1/M2/M3/M4) | `XingYiCha-macOS-arm64.dmg` | `XingYiCha-macOS-arm64.zip` |
| Intel | `XingYiCha-macOS-x86_64.dmg` | `XingYiCha-macOS-x86_64.zip` |

## 一、获取（云端构建，与 Windows 版同源）

1. 登录 GitHub（私有仓库需有权限的账号）
2. Actions 页手动 Run workflow，或打 `v*` tag 自动发布
3. Release 页面下载对应架构的 dmg 或 zip

## 二、安装与使用

**dmg 安装包**：双击 dmg → 把「星易查」拖入 Applications → 从启动台/访达打开。
首次打开若提示"无法验证开发者"（应用未签名上架 App Store，属正常）：

- 方法一：右键（或按住 Control 点击）图标 → 选「打开」→ 确认
- 方法二：系统设置 → 隐私与安全性 → 允许"星易查"

**zip 便携版**：解压出 `XingYiCha.app`，双击即可；首次打开同样按上述方法放行。

启动后自动打开浏览器访问 `http://127.0.0.1:5001`，终端窗口关闭即停止服务。

> **⚠️ 前置条件（仅当标书含老格式 `.doc`）**：`.docx` / `.pdf` / `.txt` /
> `.xlsx` / 扫描件 OCR 均开箱即用；`.doc` 的**正文提取**需先安装
> LibreOffice（`brew install --cask libreoffice`，或
> [官网](https://www.libreoffice.org/) 下载 pkg 安装）。未安装时 `.doc`
> 正文自动跳过，其元数据比对及全部其他格式、维度不受影响。

## 三、功能说明

| 功能 | 表现 |
|------|------|
| .docx / .pdf / .txt / .xlsx / 扫描件 OCR | ✅ 完整支持 |
| .doc 元数据比对 | ✅ 完整支持 |
| .doc 正文 | ⚠️ 需安装 LibreOffice（`brew install --cask libreoffice` 或官网下载），未装时自动跳过，其余不受影响 |

数据目录：`~/Library/Application Support/星易查/`（history 与 uploads）。

## 四、常见问题

- **提示"已损坏，无法打开"**：多为 Gatekeeper 对未签名应用的拦截，执行
  `xattr -cr /Applications/星易查.app` 后重试。
- **端口占用**：自动改用 5002-5010，以终端打印的实际地址为准。
- **M 系列跑 x86_64 版**：Rosetta 可运行但慢，请优先下载 arm64 版。
