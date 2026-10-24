#!/bin/bash
# MAIOS オリエン用の一括セットアップ（Mac）
#
# オリエンの場で、ターミナルに次の1行を貼り付けて実行する（管理者のパスワードは要らない）：
#   t=$(mktemp) && curl -fsSL https://raw.githubusercontent.com/AI-Management-Group/maios-setup/main/mac.sh -o "$t" && { [ "$(tail -n 1 "$t")" = "# maios-setup:end" ] || { echo "取得が途中で切れました。もう一度貼ってください" >&2; false; }; } && bash "$t"; s=$?; rm -f "$t"; (exit $s)
#   取得物は、その回だけの自分専用の一時ファイル（mktemp）に置き、最後の行が終わりの印（# maios-setup:end）だと確かめてから
#   そのファイルを動かす。終わったら消し、終了の値はそのまま返す（取得に失敗した・途中で切れた時は動かさず、0 以外で終わる。
#   同時に何人・何回走らせても取り違えない）。終わりの印は版が変わっても変えない（変えると、配った1行が全員止まる）。
#   本体は終わりの印の直前の行で呼ぶ関数にしてあり、| bash で途中まで渡された時も、関数の途中で切れていれば何もせず 0 以外で終わる
#
# やること（何度実行しても同じ結果になる）：
#   1. uv（Python を入れる道具）を ~/.local/bin に入れる
#   2. MAIOS 専用の Python 3.12 の環境を ~/.local/share/maios/python に作る（MAIOS のフックはここを最初に拾う）。
#      既にある Python（~/.local/bin/python3・Homebrew など）の中身は変えない。~/.local/bin/python3 が空いている時だけ、
#      専用の環境を呼ぶ小さな入口を置く（ターミナルから python3 で使える）。Claude の会話の中では、フックが専用の環境を先に使わせる
#   3. スキルが使う Python の部品を、専用の環境に入れる（同梱スキルのスクリプトが読み込むもの全部。一覧は下の PKGS。
#      動画・文字認識など、別のソフトも要る部品は入れない＝tools/setup/README.md の「条件付きの部品」）
#   4. Node.js（JavaScript のプログラムを動かすソフト）と、その部品2つ（pptxgenjs playwright）を入れる
#      リッチなデックを組む経路と、ブラウザで表示を測るスクリプトが使う。Node.js 20 以上が既にあればそれを使い、
#      無ければ ~/.local/share/maios/node に入れる（公式の配布物を、チェックサムで中身を確かめてから別の場所に展開し、
#      動くと確かめてから入れ替える。途中で失敗しても前の物は消えない）。~/.local/bin の node・npm・npx は空いている時だけリンクを置く。
#      部品は ~/.local/share/maios/node_parts に入れる（既にある Node.js の中身には触れない）。
#      どのフォルダに置いたスクリプトからも部品が見つかるよう、~/.node_modules をここへのリンクにする（何も無い時だけ。既にあれば、リンクでも触らない）
#   5. 画面の確認に使うブラウザ（Playwright の Chromium）を入れる。普段使いの Chrome とは別物で、スクリプトだけが使う
#   6. ターミナルから python3・node が使えるよう ~/.zprofile と ~/.zshrc に行を足す
#   7. Claude アプリと Google ドライブ（パソコン版）があるかを確かめる。どちらもこのスクリプトでは入れない（事前案内で各自が入れる）
#      Google ドライブは必須＝無ければ ❌。Claude アプリはこの1行を走らせている間に入れてもよいので、無くても ⚠️ にとどめる
# 最後に ✅／❌ の一覧を出す。記録は ~/Library/Logs/maios_setup.log。
# 作り直してよいのは ~/.local/share/maios の中だけ。それ以外の既にあるファイル・リンク・印（EXTERNALLY-MANAGED）は変えない。
# その置き場（とその上の ~/.local・~/.local/share、~/.local/bin、中の python・node・node_parts と退避）が、別の場所へのリンク・
# 自分の物でない・フォルダでない時は、どこにも書かずに止める（リンクの先にある他の物を、実行・作成・削除しないため）。
#
# 環境変数（AMG・テスト用）：MAIOS_SETUP_SKIP_APPS=1 でアプリの確認を省く／MAIOS_SETUP_FORCE_NODE=1 で既存の Node.js があっても入れる
#   MAIOS_SETUP_BASE_PYTHONS で、uv が使えない時に専用の環境の元にする Python の候補を差し替える（テスト用）
set -u
# 最後の行まで読めた時だけ本体が動く（本体の最初で MAIOS_SETUP_WHOLE=1 にする）。途中で切れたファイルは何もせず 1 で終わる
MAIOS_SETUP_WHOLE=""
trap '[ -n "$MAIOS_SETUP_WHOLE" ] || { printf "\nセットアップのファイルを最後まで取れませんでした。何も変えていません。もう一度同じ1行を貼ってください。\n" >&2; exit 1; }' EXIT
maios_setup_main() {
MAIOS_SETUP_WHOLE=1

# PKGS を変えたら IMPORTS と Windows 版（maios_setup_win.ps1 の $Pkgs・$Imports）も同じにする（tests/test_scripts_portable.py が照らす）
PKGS="python-pptx python-docx openpyxl pandas PyYAML jpholiday Pillow playwright reportlab numpy matplotlib seaborn tabulate beautifulsoup4 lxml requests"
IMPORTS="import pptx, docx, openpyxl, pandas, yaml, jpholiday, PIL, playwright, reportlab, numpy, matplotlib, seaborn, tabulate, bs4, lxml, requests"
NODE_VER="v24.21.0"
NODE_SHA_ARM64="bed7eea5325e1108f32ce5228ddd6a5f0f08a499ee42aa7442aea583702f6057"
NODE_SHA_X64="1462cb3b3046b815cf8ea436d3da450ec1a9f11dac7e5a46b0ada5305d7e8097"
NODE_PKGS="pptxgenjs playwright"
NODE_REQUIRES="require('pptxgenjs'); require('playwright')"
MAIOS_HOME="$HOME/.local/share/maios"
NODE_HOME="$MAIOS_HOME/node"
NODE_PARTS="$MAIOS_HOME/node_parts"
VENV="$MAIOS_HOME/python"
BIN="$HOME/.local/bin"
PY="$VENV/bin/python3"
WRAP_MARK="# MAIOS-SETUP: maios-python"
VENV_MARK=".maios-setup"  # このセットアップが作った専用の環境の印。印の無い物（別の環境へのリンクを含む）は使わず、消さない
LOG="$HOME/Library/Logs/maios_setup.log"
mkdir -p "$(dirname "$LOG")" 2>/dev/null
: >>"$LOG" 2>/dev/null || LOG=/dev/null
export PATH="$BIN:$PATH"

log() { printf '%s %s\n' "$(date '+%F %T')" "$*" >>"$LOG"; }
step() { printf '\n▶ %s\n' "$*"; log "== $*"; }
run() { log "\$ $*"; "$@" >>"$LOG" 2>&1; }
# 置き場所が空いている時か、もともと MAIOS の物（~/.local/share/maios を指すリンク）の時だけリンクを置く。$1＝リンク先、$2＝置く場所
link_if_free() {
  if [ -L "$2" ]; then
    case "$(readlink "$2")" in "$MAIOS_HOME"/*) ;; *) log "$2 は別の場所へのリンクなので触らない"; return 1 ;; esac
  elif [ -e "$2" ]; then
    log "$2 は既にあるので触らない"; return 1
  fi
  ln -sfn "$1" "$2"
}
# 管理する置き場の境界。リンク・自分の物でない物・フォルダでない物があれば、その場所を返す（無ければ 1）
boundary_problem() {
  for d in "$HOME/.local" "$HOME/.local/share" "$MAIOS_HOME" "$BIN"; do
    if [ -L "$d" ]; then echo "${d}（別の場所へのリンク）"; return 0; fi
    if [ -e "$d" ] && { [ ! -d "$d" ] || [ ! -O "$d" ]; }; then echo "${d}（フォルダでないか、自分の物でない）"; return 0; fi
  done
  for d in "$VENV" "$VENV.old" "$NODE_HOME" "$NODE_HOME.old" "$NODE_HOME.new" "$NODE_PARTS" "$NODE_PARTS/node_modules"; do
    if [ -L "$d" ]; then echo "${d}（別の場所へのリンク）"; return 0; fi
    if { [ -e "$d" ] && [ ! -O "$d" ]; }; then echo "${d}（自分の物でない）"; return 0; fi
  done
  # 作った後の本当の場所も、自分の HOME の下の決まった場所か確かめる（途中のどこかがリンクなら食い違う）
  if [ -d "$MAIOS_HOME" ] && [ -d "$BIN" ]; then
    home_real="$(cd -P "$HOME" 2>/dev/null && pwd -P)"
    [ "$(cd -P "$MAIOS_HOME" 2>/dev/null && pwd -P)" = "$home_real/.local/share/maios" ] || { echo "${MAIOS_HOME}（本当の場所が HOME の外）"; return 0; }
    [ "$(cd -P "$BIN" 2>/dev/null && pwd -P)" = "$home_real/.local/bin" ] || { echo "${BIN}（本当の場所が HOME の外）"; return 0; }
  fi
  return 1
}
boundary_stop() {
  log "境界の外を指している: $1 → 何も変えずに止める"
  printf '\n ❌ 次の場所が別の場所へのリンクなどになっているため、何も変えずに止めました：\n    %s\n' "$1"
  printf ' この画面の写真を撮って、Teams の MAIOS導入相談部屋へ送ってください。\n'
  exit 1
}
# 作る前に1回、作った後にもう1回確かめる（作る前に見ないと、リンクの先にフォルダを作ってしまう）
BAD="$(boundary_problem)" && boundary_stop "$BAD"
mkdir -p "$BIN" "$MAIOS_HOME" 2>/dev/null
BAD="$(boundary_problem)" && boundary_stop "$BAD"

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

# 2. Python（MAIOS 専用の環境。既にある Python は元にするだけで、中身は変えない）
step "2/6 Python 3.12"
venv_owned() { [ -d "$1" ] && [ ! -L "$1" ] && [ -f "$1/$VENV_MARK" ] && [ -f "$1/pyvenv.cfg" ]; }
py_works() { [ -x "$1/bin/python3" ] && "$1/bin/python3" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)' >/dev/null 2>&1; }
py_ok() { venv_owned "$VENV" && py_works "$VENV"; }
base_python() {
  # uv が使えない時に、専用の環境の元にする Python（3.9 以上）。Command Line Tools の無い /usr/bin/python3 は試さない
  # shellcheck disable=SC2086
  for c in "$(command -v python3 2>/dev/null || true)" ${MAIOS_SETUP_BASE_PYTHONS:-/opt/homebrew/bin/python3 /usr/local/bin/python3 /usr/bin/python3}; do
    [ -n "$c" ] && [ -x "$c" ] || continue
    if [ "$c" = /usr/bin/python3 ]; then xcode-select -p >/dev/null 2>&1 || continue; fi
    "$c" -c 'import sys, venv; sys.exit(0 if sys.version_info >= (3, 9) else 1)' >/dev/null 2>&1 && { echo "$c"; return 0; }
  done
  return 1
}
# 前の回が作り直しの途中（前の環境を python.old へ寄せた直後）で止まり、現行の場所が空なら、まず退避を戻して前の状態にする
# （動かない物でも戻す。作り直しはこの後の手順で、作れた時だけ入れ替わる）
if [ ! -e "$VENV" ] && [ ! -L "$VENV" ] && venv_owned "$VENV.old"; then
  mv "$VENV.old" "$VENV" && log "python: 途中で止まった回の退避を戻した"
fi
PY_FOREIGN=""
if { [ -e "$VENV" ] || [ -L "$VENV" ]; } && ! venv_owned "$VENV"; then
  # このセットアップが作っていない物（別の環境へのリンク・手で置いた物）は、使わず・消さず・部品も入れない
  PY_FOREIGN=1
  log "$VENV はこのセットアップが作った物ではないので触らない（$(cd -P "$VENV" 2>/dev/null && pwd -P || echo 不明)）"
elif ! py_ok; then
  # 壊れた専用の環境は脇へ寄せてから作り直し、作れなかった時は元に戻す
  OLDV=""
  if [ -e "$VENV" ] || [ -L "$VENV" ]; then OLDV="$VENV.old"; rm -rf "$OLDV"; mv "$VENV" "$OLDV"; fi
  mkdir -p "$MAIOS_HOME"
  if [ -n "$UV" ] && [ -x "$UV" ]; then
    # uv が管理する Python 3.12 だけを使う（Homebrew の更新で壊れない）。~/.local/bin には何も置かせない
    run env UV_PYTHON_PREFERENCE=only-managed UV_PYTHON_INSTALL_BIN=0 "$UV" venv --seed --python 3.12 "$VENV" \
      && [ -f "$VENV/pyvenv.cfg" ] && : >"$VENV/$VENV_MARK"
  fi
  if ! py_ok; then
    rm -rf "$VENV"
    BASE="$(base_python || true)"
    [ -n "$BASE" ] && run "$BASE" -m venv "$VENV" && [ -f "$VENV/pyvenv.cfg" ] && : >"$VENV/$VENV_MARK"
  fi
  if py_ok; then
    [ -n "$OLDV" ] && rm -rf "$OLDV"
  else
    rm -rf "$VENV"
    [ -n "$OLDV" ] && mv "$OLDV" "$VENV"
  fi
fi
if py_ok; then
  # ターミナルから python3 で使えるよう、~/.local/bin/python3 が空いている時だけ入口を置く（リンクだと専用の環境の部品が見えないので、小さなスクリプト）
  if [ -e "$BIN/python3" ] || [ -L "$BIN/python3" ]; then
    if grep -qsF "$WRAP_MARK" "$BIN/python3"; then :; else log "$BIN/python3 は既にあるので触らない（専用の環境は ${PY}）"; fi
  else
    printf '#!/bin/sh\n%s\nexec "%s" "$@"\n' "$WRAP_MARK" "$PY" >"$BIN/python3" && chmod 755 "$BIN/python3"
  fi
  R_PY="✅"
  echo "  OK（$("$PY" -c 'import platform; print(platform.python_version())' 2>/dev/null)）"
elif [ -n "$PY_FOREIGN" ]; then
  echo "  $VENV に、このセットアップが作っていない物があります。中身は変えていません。この画面の写真を AMG へ送ってください"
else
  echo "  入れられませんでした"
fi

# 3. スキルが使う Python の部品
step "3/6 スキルが使う部品"
if [ "$R_PY" = "✅" ]; then
  # shellcheck disable=SC2086
  # 入れる先は専用の環境だけ（--system は使わない）。uv で入らなければ、専用の環境の pip でやり直す
  if [ -n "$UV" ] && [ -x "$UV" ]; then
    run "$UV" pip install --python "$PY" $PKGS || run "$PY" -m pip install --disable-pip-version-check $PKGS
  else
    run "$PY" -m pip install --disable-pip-version-check $PKGS
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
# 前の回が入れ替えの途中（前の物を node.old へ寄せた直後）で止まり、現行の場所が空なら、ダウンロードの前に退避を戻して前の状態にする
# （古くて動かない物でも戻す。新しい物は、この後で展開して動くと確かめた時だけ入れ替わる）
if [ ! -e "$NODE_HOME" ] && [ ! -L "$NODE_HOME" ] && [ -d "$NODE_HOME.old" ] && [ ! -L "$NODE_HOME.old" ]; then
  mv "$NODE_HOME.old" "$NODE_HOME" && log "node: 途中で止まった回の退避を戻した"
fi
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
      # 隣の場所に展開して動くと確かめてから入れ替える（展開の途中で失敗しても、前の Node.js は残る）
      NNEW="$NODE_HOME.new"
      rm -rf "$NNEW"
      mkdir -p "$NNEW"
      if run tar -xzf "$NTMP/node.tar.gz" -C "$NNEW" --strip-components 1 && node_ok "$NNEW/bin/node"; then
        NOLD=""
        if [ -e "$NODE_HOME" ]; then NOLD="$NODE_HOME.old"; rm -rf "$NOLD"; mv "$NODE_HOME" "$NOLD"; fi
        # 動く新しい物が現行の場所に入ったと確かめるまで、前の物は消さない
        if mv "$NNEW" "$NODE_HOME" && node_ok "$NODE_HOME/bin/node"; then
          [ -n "$NOLD" ] && rm -rf "$NOLD"
        else
          log "node: 入れ替えられなかったので前の物に戻す"
          if [ -n "$NOLD" ] && [ -e "$NOLD" ]; then rm -rf "$NODE_HOME"; mv "$NOLD" "$NODE_HOME"; fi
          log "node: 戻した後 現行=$([ -x "$NODE_HOME/bin/node" ] && echo あり || echo なし)"
        fi
      else
        log "node: 展開した物が動かないので入れ替えない"
      fi
      rm -rf "$NNEW"
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
    for b in node npm npx; do link_if_free "$NODE_HOME/bin/$b" "$BIN/$b"; done
  fi
  # Node.js の playwright は、Python の playwright と同じ系列の版にそろえる（同じブラウザを共用できる）
  PWV=""
  [ "$R_PY" = "✅" ] && PWV="$("$PY" -c 'from importlib.metadata import version; print(".".join(version("playwright").split(".")[:2]))' 2>/dev/null || true)"
  NPKGS="$NODE_PKGS"
  [ -n "$PWV" ] && NPKGS="pptxgenjs playwright@~$PWV.0"
  mkdir -p "$NODE_PARTS"
  npm_parts() { PATH="$NDIR:$PATH" run npm install --prefix "$NODE_PARTS" --no-audit --no-fund --loglevel=error "$@"; }
  # shellcheck disable=SC2086
  npm_parts $NPKGS || { log "retry: npm"; sleep 5; npm_parts $NODE_PKGS; }
  if NODE_PATH="$NODE_PARTS/node_modules" "$NODE" -e "$NODE_REQUIRES" >>"$LOG" 2>&1; then
    # 次に開くターミナルとスキルが実際に使う node・npm・npx でも確かめる（PATH の先頭は ~/.local/bin。
    # 既にある物は守るので、そこに古い・壊れた物があると、MAIOS が入れた物ではなくそれが使われる）
    EFF_BAD=""
    for b in node npm npx; do
      e="$(command -v "$b" 2>/dev/null || true)"
      if [ -z "$e" ]; then EFF_BAD="$EFF_BAD ${b}（見つからない）"; continue; fi
      if [ "$b" = node ]; then
        { "$e" -e 'process.exit(parseInt(process.versions.node) >= 20 ? 0 : 1)' \
          && NODE_PATH="$NODE_PARTS/node_modules${NODE_PATH:+:$NODE_PATH}" "$e" -e "$NODE_REQUIRES"; } >>"$LOG" 2>&1 \
          || EFF_BAD="$EFF_BAD $e"
      else
        "$e" -v >>"$LOG" 2>&1 || EFF_BAD="$EFF_BAD $e"
      fi
    done
    if [ -z "$EFF_BAD" ]; then
      R_NODE="✅"
      echo "  OK（$(node -v 2>/dev/null)）"
    else
      log "node: 次に使われる物が動かない:$EFF_BAD"
      echo "  Node.js は入りましたが、先に見つかる次の物が古いか動きません（MAIOS では触りません）："
      for e in $EFF_BAD; do echo "    $e"; done
      echo "  この画面の写真を AMG へ送ってください"
    fi
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
  # 何も無い時だけ置く。既にある物（別の場所へのリンクを含む）は利用者の物として触らない（部品はターミナルの NODE_PATH で見つかる）
  if [ -L "$HOME/.node_modules" ]; then
    [ "$(readlink "$HOME/.node_modules")" = "$NODE_PARTS/node_modules" ] || log "~/.node_modules は別の場所へのリンクなので触らない"
  elif [ -e "$HOME/.node_modules" ]; then
    log "~/.node_modules は既にあるので触らない"
  else
    ln -s "$NODE_PARTS/node_modules" "$HOME/.node_modules"
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
printf ' %s スキルが使う部品\n' "$R_PKG"
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
printf ' ❌ があります。もう一度同じ1行を貼って試してください。それでも残る時は、この画面の写真を撮って AMG へ Teams で送ってください。\n'
exit 1
}
maios_setup_main "$@"
# maios-setup:end
