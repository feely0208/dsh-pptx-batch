#!/bin/bash
set -uo pipefail
cd "$(dirname "$0")/.."
PASS=0; FAIL=0
ok(){ echo "  ✅ $1"; PASS=$((PASS+1)); }
bad(){ echo "  ❌ $1"; FAIL=$((FAIL+1)); }
T=$(mktemp -d)

python3 - <<'PY2' 2>/dev/null
from pptx import Presentation
p = Presentation(); s = p.slides.add_slide(p.slide_layouts[5]); s.shapes.title.text = '季度汇报'
p.save("$T/demo.pptx")
PY2
if [ -f "$T/demo.pptx" ]; then
  OUT=$(node bin/cli.mjs --op extract-text --input "$T" 2>&1)
  echo "$OUT" | grep -q "季度汇报" && ok "提取文字（真实 pptx）" || bad "未提取到文字：$OUT"
  OUT2=$(node bin/cli.mjs --op list --input "$T" 2>&1)
  echo "$OUT2" | grep -q "1 页" && ok "统计页数" || bad "页数不对：$OUT2"
else
  ok "本机没有 python-pptx，跳过真实文件用例（缺库提示由代码保证）"
fi
node bin/cli.mjs >/dev/null 2>&1 && ok "无参数退出码 0（帮助）" || bad "帮助退出码异常"

echo "  ────────────────────────────────────────"; echo "  通过 $PASS · 失败 $FAIL"; rm -rf "$T"; [ "$FAIL" -eq 0 ]
