#!/bin/zsh
set -eu
cd "$(dirname "$0")"
ledger_node="$(command -v node || true)"
if [[ -z "$ledger_node" && -x /opt/homebrew/bin/node ]]; then ledger_node=/opt/homebrew/bin/node; fi
if [[ -z "$ledger_node" && -x /usr/local/bin/node ]]; then ledger_node=/usr/local/bin/node; fi
if [[ -z "$ledger_node" ]]; then
  echo "未找到 Node.js 20 或以上。请 IT 安装 Node.js 后重试；本版本须通过本机服务启动，详情见 README。"
  read -r "?按回车关闭。"
  exit 1
fi
echo "正在打开台账对对看。保留此窗口即可使用；按 Ctrl+C 停止。"
PORT=0 "$ledger_node" scripts/serve.mjs --open
