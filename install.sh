#!/usr/bin/env bash
# bel-hot-sales installer (macOS / Linux)
#
# Copies the skill into ~/.agents/skills/bel-hot-sales and checks dependencies.
# Re-runnable: safe to run again to update.
#
#   bash install.sh

set -euo pipefail

SKILL_NAME="bel-hot-sales"
SKILLS_ROOT="${SKILLS_ROOT:-$HOME/.agents/skills}"
SOURCE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
TARGET="$SKILLS_ROOT/$SKILL_NAME"
SKIP_DEPS="${SKIP_DEPS:-}"

c_ok()   { printf '  \033[32m[OK]\033[0m %s\n' "$1"; }
c_warn() { printf '  \033[33m[!]\033[0m  %s\n' "$1"; }
c_fail() { printf '  \033[31m[X]\033[0m  %s\n' "$1"; }
c_info() { printf '  %s\n' "$1"; }

echo
printf '\033[36mbel-hot-sales installer\033[0m\n'
printf '\033[36m=======================\033[0m\n'

# --- 1. sanity-check the source -------------------------------------------
echo
echo "1. Checking source"
if [ ! -f "$SOURCE_DIR/SKILL.md" ]; then
  c_fail "SKILL.md not found in '$SOURCE_DIR'."
  c_info "Run this script from inside the cloned repo, or set SOURCE_DIR=<path>."
  exit 1
fi
if [ ! -f "$SOURCE_DIR/scripts/bel_pipeline.py" ]; then
  c_warn "scripts/bel_pipeline.py not found. The skill will still load, but its"
  c_warn "bundled script will be missing and the agent must reimplement the logic."
else
  c_ok "found SKILL.md and scripts/bel_pipeline.py"
fi

# --- 2. copy ---------------------------------------------------------------
echo
echo "2. Installing to $TARGET"
[ -d "$TARGET" ] && { c_warn "Target already exists; replacing."; rm -rf "$TARGET"; }
mkdir -p "$TARGET"

for f in SKILL.md README.md LICENSE; do
  [ -f "$SOURCE_DIR/$f" ] && cp "$SOURCE_DIR/$f" "$TARGET/"
done
for d in scripts evals; do
  if [ -d "$SOURCE_DIR/$d" ]; then
    cp -R "$SOURCE_DIR/$d" "$TARGET/"
    find "$TARGET/$d" -type d -name __pycache__ -exec rm -rf {} + 2>/dev/null || true
  fi
done
c_ok "copied files"

# --- 3. verify the layout skill discovery requires ------------------------
echo
echo "3. Verifying layout"
if [ ! -f "$TARGET/SKILL.md" ]; then
  c_fail "SKILL.md is not at the root of $TARGET -- the skill will NOT load."
  exit 1
fi
if head -n 1 "$TARGET/SKILL.md" | grep -q '^---'; then
  c_ok "SKILL.md present with frontmatter at the skill root"
else
  c_warn "SKILL.md does not start with '---' YAML frontmatter; it may not be parsed."
fi
declared="$(grep -m1 '^name:' "$TARGET/SKILL.md" | sed 's/^name:[[:space:]]*//' | tr -d '\r' || true)"
if [ -n "$declared" ] && [ "$declared" != "$SKILL_NAME" ]; then
  c_warn "frontmatter name is '$declared' but the folder is '$SKILL_NAME'."
  c_warn "Keep them equal to avoid confusion."
elif [ -n "$declared" ]; then
  c_ok "frontmatter name matches folder name ('$SKILL_NAME')"
fi

# --- 4. dependencies -------------------------------------------------------
echo
echo "4. Checking dependencies"
if [ -n "$SKIP_DEPS" ]; then
  c_info "skipped (SKIP_DEPS set)"
else
  PY=""
  for cand in python3 python; do
    if command -v "$cand" >/dev/null 2>&1; then PY="$cand"; break; fi
  done
  if [ -z "$PY" ]; then
    c_fail "Python not found on PATH."
    c_info "The skill is installed and will still load, but the bundled script"
    c_info "cannot run. Either install Python 3, or let the agent follow SKILL.md"
    c_info "and implement the logic itself."
  else
    c_ok "python: $(command -v "$PY")"
    if ver="$("$PY" -c 'import openpyxl;print(openpyxl.__version__)' 2>/dev/null)"; then
      c_ok "openpyxl $ver"
    else
      c_warn "openpyxl is not installed."
      printf "  Install it now with 'pip install openpyxl'? [y/N] "
      read -r ans
      case "$ans" in
        y|Y|yes|YES)
          if "$PY" -m pip install --quiet openpyxl; then c_ok "openpyxl installed"
          else c_fail "pip install failed -- install it manually."; fi
          ;;
        *) c_info "skipped. Run: $PY -m pip install openpyxl" ;;
      esac
    fi
  fi
fi

# --- 5. done ---------------------------------------------------------------
echo
printf '\033[36mDone.\033[0m\n'
c_info "Installed to: $TARGET"
echo
c_info "Test it by asking your agent something like:"
c_info "  '这个十部半托管库存表给你，筛选出两周都有销量的产品，按 V3 格式输出'"
echo
c_info "Note: a skill is only discovered when the agent starts a NEW session."
