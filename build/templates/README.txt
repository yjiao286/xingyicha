╔══════════════════════════════════════════════════════════╗
║          星易查 · 围串标风险识别分析系统                  ║
║          Linux 便携版 (离线运行)                         ║
╚══════════════════════════════════════════════════════════╝

【一、运行】
  方式 A（推荐，桌面环境如银河麒麟/Ubuntu 等）：
        首次：在文件管理器里双击 install.sh → 选“运行/在终端中运行”
              （它会把桌面/开始菜单快捷方式装好，并自动启动）
        以后：从开始菜单点“星易查”，或桌面图标（首次需右键→“允许启动”）
  方式 B（终端，最稳）：
        cd 进本目录后执行  ./run.sh
  启动后会自动打开默认浏览器访问 http://127.0.0.1:5001
  关闭启动它的终端窗口即停止服务。

  若双击 xingyicha.desktop 仍弹出文本编辑器：
        说明桌面未把它当可执行快捷方式。用方式 B（终端 ./run.sh）最省事，
        或先执行 ./install.sh 安装正式快捷方式。

【二、环境要求】
  · Linux x86_64 或 aarch64（按压缩包后缀选择）
  · glibc ≥ 2.28 （Ubuntu 20.04+ / Debian 11+ / RHEL/Rocky 8+ / Fedora 32+ 等）
  · 任意现代浏览器（用于打开前端）
  · 无需安装 Python、无需联网

【三、自检】
  首次使用或迁移后，可执行：
        ./verify.sh
  会校验内置 Python 与依赖是否完好。

【四、数据目录】
  上传的投标文件与分析历史默认存于：
        ~/.xingyicha/uploads
        ~/.xingyicha/history
  可用环境变量覆盖：
        XINGYICHA_DATA=/your/path  ./run.sh
        XINGYICHA_PORT=6001        ./run.sh   # 改端口
        XINGYICHA_HOST=0.0.0.0     ./run.sh   # 允许局域网访问（默认仅本机）

【五、文件格式】
  · .docx  原生支持
  · .pdf   原生支持
  · .txt   原生支持（UTF-8/GB18030 自动识别，OCR 标书输出可用；无文档元数据）
  · .doc   需内置 antiword（已尽力打包；若 verify.sh 提示缺失，.doc 不可用，其余格式正常）

【六、目录结构】
  python/       内置可重定位 CPython 3.11（含依赖）
  app/          应用代码与前端资源（app.py / templates / static）
  bin/antiword  .doc 文本提取工具
  run.sh        启动器（终端启动）
  install.sh    桌面/开始菜单快捷方式安装（双击启动）
  verify.sh     离线自检
  xingyicha.desktop  桌面快捷方式入口

【七、常见问题：页面提示“服务器内部错误 / trace_id / duration_ms”】
  本程序的错误返回是 {"error":"<具体原因>"}，不会带 trace_id / duration_ms。
  若你看到 {"error":"服务器内部错误","_meta":{"trace_id":...},"duration_ms":...}
  这种结构，那不是本程序返回的，而是 系统/浏览器 的 HTTP 代理拦截了
  对 127.0.0.1 / localhost 的请求并转发失败所致（政企网络、麒麟常见）。
  排查：
    1) 在浏览器里直接访问 http://127.0.0.1:5001 ，而不是 localhost；
    2) 关闭系统/浏览器代理，或把 127.0.0.1、localhost 加入“代理例外/不代理”；
       - 麒麟设置→网络→代理→“忽略主机”填 127.0.0.1,localhost
       - 或环境变量： export no_proxy="127.0.0.1,localhost"  再 ./run.sh
    3) 终端直连验证（不经代理）：
         curl --noproxy '*' http://127.0.0.1:5001/
       返回 HTML 即服务正常；返回上面那个 JSON 即确认是代理问题。
    4) 服务真实日志在 ~/.xingyicha/server.log，可查看后端是否真的报错。

