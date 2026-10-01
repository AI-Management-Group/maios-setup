#!/bin/bash
# MAIOS オリエン用の一括セットアップ（Mac）
#
# オリエンの場で、ターミナルに次の1行を貼り付けて実行する（管理者のパスワードは要らない）：
#   curl -fsSL https://raw.githubusercontent.com/AI-Management-Group/maios-setup/main/mac.sh | bash
#
# やること（何度実行しても同じ結果になる）：
#   1. uv（Python を入れる道具）を ~/.local/bin に入れる
#   2. Python 3.12 を ~/.local/bin/python3 に入れる（MAIOS のフックはここを PATH に無くても拾う）
#   3. スキルが使う部品8つを入れる（python-pptx openpyxl pandas PyYAML jpholiday Pillow playwright reportlab）
#   4. Node.js（JavaScript のプログラムを動かすソフト）と、その部品2つ（pptxgenjs playwright）を入れる
#      リッチなデックを組む経路と、ブラウザで表示を測るスクリプトが使う。Node.js 20 以上が既にあればそれを使い、
#      無ければ ~/.local/share/maios/node に入れる（公式の配布物を、チェックサムで中身を確かめてから展開する）。
#      部品は ~/.local/share/maios/node_parts に入れる（既にある Node.js の中身には触れない）。
#      どのフォルダに置いたスクリプトからも部品が見つかるよう、~/.node_modules をここへのリンクにする（無い時だけ）
#   5. 画面の確認に使うブラウザ（Playwright の Chromium）を入れる。普段使いの Chrome とは別物で、スクリプトだけが使う
#   6. ターミナルから python3・node が使えるよう ~/.zprofile と ~/.zshrc に行を足す
#   7. Claude アプリと Google ドライブ（パソコン版）があるかを確かめる。どちらもこのスクリプトでは入れない（事前案内で各自が入れる）
#      Google ドライブは必須＝無ければ ❌。Claude アプリはこの1行を走らせている間に入れてもよいので、無くても ⚠️ にとどめる
# 最後に ✅／❌ の一覧を出す。記録は ~/Library/Logs/maios_setup.log。
#
# 環境変数（AMG・テスト用）：MAIOS_SETUP_SKIP_APPS=1 でアプリの確認を省く／MAIOS_SETUP_FORCE_NODE=1 で既存の Node.js があっても入れる
set -u

PKGS="python-pptx openpyxl pandas PyYAML jpholiday Pillow playwright reportlab"
IMPORTS="import pptx, openpyxl, pandas, yaml, jpholiday, PIL, playwright, reportlab"
NODE_VER="v24.21.0"
NODE_SHA_ARM64="bed7eea5325e1108f32ce5228ddd6a5f0f08a499ee42aa7442aea583702f6057"
NODE_SHA_X64="1462cb3b3046b815cf8ea436d3da450ec1a9f11dac7e5a46b0ada5305d7e8097"
NODE_PKGS="pptxgenjs playwright"
NODE_REQUIRES="require('pptxgenjs'); require('playwright')"
NODE_HOME="$HOME/.local/share/maios/node"
NODE_PARTS="$HOME/.local/share/maios/node_parts"
BIN="$HOME/.local/bin"
PY="$BIN/python3"
LOG="$HOME/Library/Logs/maios_setup.log"
mkdir -p "$BIN" "$(dirname "$LOG")" 2>/dev/null
: >>"$LOG" 2>/dev/null || LOG=/dev/null
export PATH="$BIN:$PATH"

log() { printf '%s %s\n' "$(date '+%F %T')" "$*" >>"$LOG"; }
step() { printf '\n▶ %s\n' "$*"; log "== $*"; }
run() { log "\$ $*"; "$@" >>"$LOG" 2>&1; }

R_PY="❌" R_PKG="❌" R_NODE="❌" R_BROWSER="⚠️" R_APP="－" R_DRIVE="－"

printf '\nMAIOS のセットアップを始めます（5〜10分）。終わるまで、この画面を閉じないでください。\n'
log "start $(sw_vers -productVersion 2>/dev/null) $(uname -m)"

# 1. uv
step "1/6 Python を入れる道具（uv）"
UV="$(command -v uv 2>/dev/null || true)"
if [ -z "$UV" ]; then
  if curl -LsSf https://astral.sh/uv/install.sh 2>>"$LOG" | env UV_INSTALL_DIR="$BIN" INSTALLER_NO_MODIFY_PATH=1 sh >>"$LOG" 2>&1; then
    UV="$BIN/uv"
  fi
fi
if [ -n "$UV" ] && [ -x "$UV" ]; then echo "  OK"; else echo "  入れられませんでした（ネットワークを確認してください）"; fi

# 2. Python
step "2/6 Python 3.12"
py_ok() { [ -x "$PY" ] && "$PY" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)' >/dev/null 2>&1; }
if ! py_ok && [ -n "$UV" ] && [ -x "$UV" ]; then
  run "$UV" python install 3.12 --default
fi
if py_ok; then
  # uv の Python は「外から部品を足さない」印（EXTERNALLY-MANAGED）が付いている。MAIOS 専用の Python なので外して、スキルの pip も通るようにする
  MARK="$("$PY" -c 'import os, sysconfig; print(os.path.join(sysconfig.get_path("stdlib"), "EXTERNALLY-MANAGED"))' 2>/dev/null)"
  [ -n "$MARK" ] && [ -f "$MARK" ] && run rm -f "$MARK"
  R_PY="✅"
  echo "  OK（$("$PY" -c 'import platform; print(platform.python_version())' 2>/dev/null)）"
else
  echo "  入れられませんでした"
fi

# 3. 部品8つ
step "3/6 スキルが使う部品（8つ）"
if [ "$R_PY" = "✅" ]; then
  # shellcheck disable=SC2086
  if [ -n "$UV" ] && [ -x "$UV" ]; then
    run "$UV" pip install --python "$PY" --system --break-system-packages $PKGS
  else
    run "$PY" -m pip install --disable-pip-version-check --break-system-packages $PKGS
  fi
  if "$PY" -c "$IMPORTS" >>"$LOG" 2>&1; then R_PKG="✅"; echo "  OK"; else echo "  一部が入りませんでした"; fi
else
  echo "  Python が無いので飛ばしました"
fi

# 4. Node.js と部品2つ
step "4/6 Node.js と部品（2つ）"
node_ok() {
  # 20 以上で、隣に npm（部品を入れる道具）があれば使える
  [ -n "$1" ] && [ -x "$1" ] && [ -x "$(dirname "$1")/npm" ] || return 1
  "$1" -e 'process.exit(parseInt(process.versions.node) >= 20 ? 0 : 1)' >/dev/null 2>&1
}
find_node() {
  node_ok "$NODE_HOME/bin/node" && { echo "$NODE_HOME/bin/node"; return 0; }
  [ -n "${MAIOS_SETUP_FORCE_NODE:-}" ] && return 1
  for c in "$(command -v node 2>/dev/null || true)" /opt/homebrew/bin/node /usr/local/bin/node; do
    node_ok "$c" && { echo "$c"; return 0; }
  done
  return 1
}
NODE="$(find_node || true)"
if [ -z "$NODE" ]; then
  case "$(uname -m)" in
    arm64) NARCH="arm64"; NSHA="$NODE_SHA_ARM64" ;;
    *) NARCH="x64"; NSHA="$NODE_SHA_X64" ;;
  esac
  NTMP="$(mktemp -d)"
  echo "  ダウンロードしています…"
  if curl -fsSL "https://nodejs.org/dist/$NODE_VER/node-$NODE_VER-darwin-$NARCH.tar.gz" -o "$NTMP/node.tar.gz" 2>>"$LOG"; then
    if [ "$(shasum -a 256 "$NTMP/node.tar.gz" 2>/dev/null | cut -d' ' -f1)" = "$NSHA" ]; then
      rm -rf "$NODE_HOME"
      mkdir -p "$NODE_HOME"
      run tar -xzf "$NTMP/node.tar.gz" -C "$NODE_HOME" --strip-components 1
    else
      log "node: checksum mismatch"
    fi
  else
    log "node: download failed"
  fi
  rm -rf "$NTMP"
  NODE="$(find_node || true)"
fi
if [ -n "$NODE" ]; then
  NDIR="$(dirname "$NODE")"
  if [ "$NDIR" = "$NODE_HOME/bin" ]; then
    for b in node npm npx; do ln -sf "$NODE_HOME/bin/$b" "$BIN/$b"; done
  fi
  # Node.js の playwright は、Python の playwright と同じ系列の版にそろえる（同じブラウザを共用できる）
  PWV="$("$PY" -c 'from importlib.metadata import version; print(".".join(version("playwright").split(".")[:2]))' 2>/dev/null || true)"
  NPKGS="$NODE_PKGS"
  [ -n "$PWV" ] && NPKGS="pptxgenjs playwright@~$PWV.0"
  mkdir -p "$NODE_PARTS"
  npm_parts() { PATH="$NDIR:$PATH" run npm install --prefix "$NODE_PARTS" --no-audit --no-fund --loglevel=error "$@"; }
  # shellcheck disable=SC2086
  npm_parts $NPKGS || { log "retry: npm"; sleep 5; npm_parts $NODE_PKGS; }
  if NODE_PATH="$NODE_PARTS/node_modules" "$NODE" -e "$NODE_REQUIRES" >>"$LOG" 2>&1; then
    R_NODE="✅"
    echo "  OK（$("$NODE" -v 2>/dev/null)）"
  else
    echo "  部品が入りませんでした"
  fi
else
  echo "  入れられませんでした（ネットワークを確認してください）"
fi

# 5. 画面の確認に使うブラウザ（Playwright の Chromium）。無くても Chrome で代わりが利くスキルが多いので、入らなくても ⚠️ にとどめる
step "5/6 画面の確認に使うブラウザ"
PW_PY='from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    p.chromium.launch().close()'
PW_JS="require('playwright').chromium.launch().then(b => b.close()).catch(e => { console.error(String(e)); process.exit(1); })"
B_PY="" B_JS=""
if [ "$R_PKG" = "✅" ]; then
  "$PY" -c "$PW_PY" >/dev/null 2>&1 || { echo "  ダウンロードしています…（1〜3分）"; run "$PY" -m playwright install chromium; }
  "$PY" -c "$PW_PY" >>"$LOG" 2>&1 && B_PY=1
fi
if [ "$R_NODE" = "✅" ]; then
  node_browser() { NODE_PATH="$NODE_PARTS/node_modules" "$NODE" -e "$PW_JS"; }
  node_browser >/dev/null 2>&1 || PATH="$NDIR:$PATH" run "$NODE" "$NODE_PARTS/node_modules/playwright/cli.js" install chromium
  node_browser >>"$LOG" 2>&1 && B_JS=1
fi
if [ -n "$B_PY" ] && [ -n "$B_JS" ]; then R_BROWSER="✅"; echo "  OK"; else log "browser py=${B_PY:-0} js=${B_JS:-0}"; echo "  入りませんでした（後からでも入れられます）"; fi

# 6. ターミナルの設定（新しく開くターミナルから python3・node が使え、Node.js の部品が見つかる）
step "6/6 仕上げ"
LINE='export PATH="$HOME/.local/bin:$PATH"  # MAIOS'
NLINE='case ":${NODE_PATH:-}:" in *"/maios/node_parts/node_modules:"*) ;; *) export NODE_PATH="$HOME/.local/share/maios/node_parts/node_modules${NODE_PATH:+:$NODE_PATH}" ;; esac  # MAIOS'
for rc in "$HOME/.zprofile" "$HOME/.zshrc"; do
  grep -qs 'export PATH="$HOME/.local/bin:$PATH"' "$rc" || printf '\n%s\n' "$LINE" >>"$rc"
  if [ "$R_NODE" = "✅" ]; then
    grep -qsF 'maios/node_parts/node_modules' "$rc" || printf '%s\n' "$NLINE" >>"$rc"
  fi
done
if [ "$R_NODE" = "✅" ]; then
  # Node.js は ~/.node_modules も必ず探す。ターミナルの設定が読まれない場面（アプリから起動した時など）でも部品が見つかる
  if [ -L "$HOME/.node_modules" ] || [ ! -e "$HOME/.node_modules" ]; then
    ln -sfn "$NODE_PARTS/node_modules" "$HOME/.node_modules"
  else
    log "~/.node_modules は既にあるので触らない"
  fi
fi
echo "  OK"

# アプリの確認（入れるのはオリエンの画面の手順で）
if [ -z "${MAIOS_SETUP_SKIP_APPS:-}" ]; then
  if [ -d /Applications/Claude.app ] || [ -d "$HOME/Applications/Claude.app" ]; then R_APP="✅"; else R_APP="⚠️"; fi
  if [ -d "/Applications/Google Drive.app" ] || [ -d "$HOME/Applications/Google Drive.app" ]; then R_DRIVE="✅"; else R_DRIVE="❌"; fi
fi
log "result py=$R_PY pkg=$R_PKG node=$R_NODE browser=$R_BROWSER app=$R_APP drive=$R_DRIVE python=$PY node_bin=${NODE:-}"

printf '\n==============================\n'
printf ' MAIOS セットアップの結果\n'
printf '==============================\n'
printf ' %s Python\n' "$R_PY"
printf ' %s 部品8つ\n' "$R_PKG"
printf ' %s Node.js と部品2つ\n' "$R_NODE"
printf ' %s 画面確認用のブラウザ\n' "$R_BROWSER"
[ "$R_APP" = "－" ] || printf ' %s Claude アプリ\n' "$R_APP"
[ "$R_DRIVE" = "－" ] || printf ' %s Google ドライブ\n' "$R_DRIVE"
printf '==============================\n'
if [ "$R_PY" = "✅" ] && [ "$R_PKG" = "✅" ] && [ "$R_NODE" = "✅" ] && [ "$R_DRIVE" != "❌" ]; then
  printf ' 完了です。Claude アプリを一度終了して開き直し、次の手順へ進んでください。\n'
  [ "$R_APP" = "⚠️" ] && printf ' Claude アプリがまだ見つかりません。入れていなければ、画面の手順で入れてください。\n'
  [ "$R_BROWSER" = "⚠️" ] && printf ' 画面確認用のブラウザが入りませんでした。このままでも始められます。後でもう一度この1行を貼ると入ります。\n'
  exit 0
fi
[ "$R_DRIVE" = "❌" ] && printf ' Google ドライブ（パソコン版）が入っていません。事前案内の手順で入れてから、もう一度この1行を貼ってください。\n'
printf ' ❌ があります。手を挙げて、この画面を AMG に見せてください。\n'
exit 1
