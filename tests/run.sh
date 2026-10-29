#!/bin/bash
# @deepwhale-cn/dsh-pptx-batch · 双向验证
# 正例：用 python-pptx 现场造 pptx（占位符 / 表格 / 图片 / 组合形状），
#       跑 fill / extract-text / replace-image / info，逐项断言产物确实变了；
# 反例：坏输入、缺参数、未知子命令必须非 0 退出并说明原因。
# 全程前后比对输入文件 sha256，确认没有改动输入。
set -uo pipefail
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"
cd "$(dirname "$0")/.."

PASS=0; FAIL=0
ok(){ echo "  ✅ $1"; PASS=$((PASS+1)); }
bad(){ echo "  ❌ $1"; FAIL=$((FAIL+1)); }
check(){ # check <描述> <命令...>
  local desc="$1"; shift
  if "$@" >/dev/null 2>"$TMP/last_err"; then ok "$desc"
  else bad "$desc"; sed -n '1,3p' "$TMP/last_err" | sed 's/^/       /'; fi
}
say(){ echo ""; echo "  【$1】"; }

TMP="$(mktemp -d "${TMPDIR:-/tmp}/dsh-pptx-batch-test.XXXXXX")"
trap 'rm -rf "$TMP"' EXIT
FIX="$TMP/输入/模板"
OUT="$TMP/输出"
mkdir -p "$FIX"

echo ""; echo "  @deepwhale-cn/dsh-pptx-batch · 双向验证"
echo "  ────────────────────────────────────────────────────────"

# ── 造正例样本（python-pptx，不依赖手工二进制） ────────────────────
python3 - "$FIX" <<'PY' || { echo "  ❌ 造样例失败"; exit 1; }
import struct, sys, zlib
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt

root = Path(sys.argv[1])


def png(path, rgb, size=16):
    """手写纯色 PNG（不依赖 Pillow）。"""
    w = h = size
    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))

    def chunk(tag, data):
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)
    path.write_bytes(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
                     + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


png(root / "old-logo.png", (200, 30, 30))
png(root / "new-logo.png", (20, 120, 220))

prs = Presentation()
s1 = prs.slides.add_slide(prs.slide_layouts[5])          # Title Only
s1.shapes.title.text = "{{title}}"

box = s1.shapes.add_textbox(Inches(0.6), Inches(1.6), Inches(6.5), Inches(1.8))
frame = box.text_frame
frame.word_wrap = True
p1 = frame.paragraphs[0]
r1 = p1.add_run(); r1.text = "客户："; r1.font.bold = True
r2 = p1.add_run(); r2.text = "{{client}}"; r2.font.italic = True
r3 = p1.add_run(); r3.text = "（已确认）"
p2 = frame.add_paragraph()
a = p2.add_run(); a.text = "{{am"; a.font.bold = True; a.font.size = Pt(24)   # 跨 run 占位符
b = p2.add_run(); b.text = "ount}}"
frame.add_paragraph().text = "日期：{{date}}"

table = s1.shapes.add_table(2, 3, Inches(0.6), Inches(3.3), Inches(7), Inches(1.2)).table
table.cell(0, 0).text = "项目"; table.cell(0, 1).text = "客户"; table.cell(0, 2).text = "金额"
table.cell(1, 0).text = "报价"; table.cell(1, 1).text = "{{client}}"; table.cell(1, 2).text = "{{amount}}"

pic = s1.shapes.add_picture(str(root / "old-logo.png"), Inches(9), Inches(0.8), Inches(1.5), Inches(1.5))
pic.name = "Logo"

s2 = prs.slides.add_slide(prs.slide_layouts[6])          # Blank
note = s2.shapes.add_textbox(Inches(0.6), Inches(0.6), Inches(6), Inches(1.2))
note.text_frame.text = "第二页备注：{{client}}"
named = s2.shapes.add_textbox(Inches(0.6), Inches(2.2), Inches(4), Inches(0.8))
named.name = "说明"
named.text_frame.text = "这是说明形状，不是图片"

group = s2.shapes.add_group_shape()
gbox = group.shapes.add_textbox(Inches(0.6), Inches(4.0), Inches(4), Inches(0.8))
gbox.text_frame.text = "组合形状内：{{client}}"

prs.save(str(root / "报价单.pptx"))

prs2 = Presentation()
t = prs2.slides.add_slide(prs2.slide_layouts[5])
t.shapes.title.text = "附件"
t.shapes.add_textbox(Inches(0.6), Inches(1.6), Inches(6), Inches(1.2)).text_frame.text = "{{client}} 的附件"
prs2.save(str(root / "附件.pptx"))

# PowerPoint 打开时产生的临时文件：必须跳过，不能当成输入
(root / "~$临时.pptx").write_bytes(b"not a real pptx")
PY

BEFORE=$(shasum -a 256 "$FIX"/*.pptx | shasum -a 256)
ORIG_IMG_SHA=$(python3 - "$FIX/报价单.pptx" <<'PY'
import hashlib, sys
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
for slide in Presentation(sys.argv[1]).slides:
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            print(hashlib.sha256(shape.image.blob).hexdigest())
PY
)

say "CLI 形态"
node bin/cli.mjs --help >"$TMP/help.txt" 2>&1 && ok "--help 退出 0" || bad "--help 退出非 0"
grep -q "子命令" "$TMP/help.txt" && ok "--help 列出子命令" || bad "--help 缺少子命令列表"
grep -q "python-pptx" "$TMP/help.txt" && ok "--help 说明 python-pptx 依赖" || bad "--help 未说明依赖"
grep -q "replace-image" "$TMP/help.txt" && ok "--help 列出 replace-image" || bad "--help 缺 replace-image"
node bin/cli.mjs >"$TMP/noargs.txt" 2>&1; RC=$?
[ "$RC" -eq 0 ] && ok "无参数时打印 help 且退出 0" || bad "无参数未退出 0（$RC）"
grep -q "fill" "$TMP/noargs.txt" && ok "无参数输出与 --help 一致" || bad "无参数没有打印 help"
DEP=$(python3 -S lib/pptx_tool.py info '{}' 2>&1); RC=$?
[ "$RC" -eq 3 ] && ok "缺 python-pptx 依赖时退出码 3" || bad "缺依赖退出码应为 3，实际 $RC"
echo "$DEP" | grep -q "pip install python-pptx" && ok "缺依赖明确提示安装命令（不静默）" || bad "缺依赖没有安装提示"

say "正例 · fill 批量填占位符（正文 + 表格 + 组合形状 + 跨 run）"
node bin/cli.mjs fill --input "$FIX" --out "$OUT/fill" \
  --set title=季度经营简报 --set client=客户甲 --set amount=128000 --set date=2026-10-24 \
  --json >"$TMP/fill.json" 2>&1
RC=$?
[ "$RC" -eq 0 ] && ok "fill 退出 0" || bad "fill 退出 $RC"
[ "$(python3 -c "import json,sys;print(len(json.load(open(sys.argv[1]))['files']))" "$TMP/fill.json" 2>/dev/null)" = "2" ] \
  && ok "只处理 2 个 .pptx，跳过 ~\$ 临时文件" || bad "~\$ 临时文件没有被跳过，或文件数不对"
cat > "$TMP/check_fill.py" <<'PY'
import glob, json, os, sys
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt

out, json_path, orig_sha = sys.argv[1], sys.argv[2], sys.argv[3]
data = json.load(open(json_path, encoding="utf-8"))
assert data.get("ok") is True, data
assert data["action"] == "fill"
assert data["replacements"] >= 8, "替换次数过少: %s" % data["replacements"]
assert data["unfilled"] == [], "还有未赋值占位符: %s" % data["unfilled"]

files = sorted(glob.glob(os.path.join(out, "*.pptx")))
assert len(files) == 2, "期望 2 个输出文件，实际 %d" % len(files)


def iter_shapes(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(sh.shapes)


def texts(prs):
    out_texts = []
    for slide in prs.slides:
        for sh in iter_shapes(slide.shapes):
            if sh.has_text_frame:
                out_texts.append(sh.text_frame.text)
            if sh.has_table:
                for row in sh.table.rows:
                    for cell in row.cells:
                        out_texts.append(cell.text)
    return out_texts


def paragraphs(prs):
    for slide in prs.slides:
        for sh in iter_shapes(slide.shapes):
            if sh.has_text_frame:
                yield from sh.text_frame.paragraphs
            if sh.has_table:
                for row in sh.table.rows:
                    for cell in row.cells:
                        yield from cell.text_frame.paragraphs


all_text = []
for path in files:
    prs = Presentation(path)
    t = texts(prs)
    all_text.extend(t)
    joined = "\n".join(t)
    assert "{{" not in joined, "%s 里还有占位符：%r" % (os.path.basename(path), joined[:200])

    if "报价单" in path:
        assert "季度经营简报" in joined, "标题占位符没填上: %r" % joined[:200]
        assert "客户甲" in joined and "128000" in joined, "正文/表格没填上"
        assert "2026-10-24" in joined, "日期占位符没填上"
        # 段落内 run 格式：斜体 run 被替换后仍斜体，前一个 run 仍加粗
        para = next(p for p in paragraphs(prs) if p.text == "客户：客户甲（已确认）")
        assert para.runs[0].text == "客户：" and para.runs[0].font.bold is True, "替换破坏了前一段 run"
        assert para.runs[1].text == "客户甲" and para.runs[1].font.italic is True, "替换丢了斜体格式"
        # 跨 run 占位符：合并到起始 run，加粗与字号保留
        para2 = next(p for p in paragraphs(prs) if p.text == "128000")
        assert para2.runs[0].font.bold is True, "跨 run 替换后加粗丢了"
        assert para2.runs[0].font.size == Pt(24), "跨 run 替换后字号丢了"
        # 表格单元格
        rows = [[c.text for c in row.cells] for slide in prs.slides for sh in iter_shapes(slide.shapes) if sh.has_table for row in sh.table.rows]
        flat = [c for r in rows for c in r]
        assert "客户甲" in flat and "128000" in flat, "表格单元格没替换: %s" % flat
        # 图片不能被 fill 动到
        import hashlib
        pics = [sh for slide in prs.slides for sh in iter_shapes(slide.shapes) if sh.shape_type == MSO_SHAPE_TYPE.PICTURE]
        assert hashlib.sha256(pics[0].image.blob).hexdigest() == orig_sha, "fill 不应改动图片内容"

assert any("组合形状内：客户甲" in t for t in all_text), "组合形状里的占位符没替换"
PY
check "fill 后无残留 {{}}，正文/表格/组合形状/跨 run 全部替换" \
  python3 "$TMP/check_fill.py" "$OUT/fill" "$TMP/fill.json" "$ORIG_IMG_SHA"

AFTER=$(shasum -a 256 "$FIX"/*.pptx | shasum -a 256)
[ "$BEFORE" = "$AFTER" ] && ok "输入文件未被修改（sha256 一致）" || bad "输入文件被修改了"

say "正例 · extract-text 提取每页文本"
node bin/cli.mjs extract-text --input "$OUT/fill" --out "$OUT/文本" --json >"$TMP/extract.json" 2>&1
RC=$?
[ "$RC" -eq 0 ] && ok "extract-text 退出 0" || bad "extract-text 退出 $RC"
cat > "$TMP/check_extract.py" <<'PY'
import glob, json, os, sys
out, json_path = sys.argv[1], sys.argv[2]
data = json.load(open(json_path, encoding="utf-8"))
assert data.get("ok") is True, data
files = sorted(glob.glob(os.path.join(out, "*.txt")))
assert len(files) == 2, "期望 2 个 txt，实际 %d" % len(files)
main = open(next(p for p in files if "报价单" in p), encoding="utf-8").read()
assert "=== 第 1 页 · 季度经营简报 ===" in main, "页头没有标题:\n" + main[:300]
assert "客户：客户甲" in main, "正文没提取到"
assert "128000" in main, "跨 run 填上的金额没提取到"
assert "项目 | 客户 | 金额" in main, "表格表头没提取到:\n" + main[:400]
assert "报价 | 客户甲 | 128000" in main, "表格数据没提取到"
assert "=== 第 2 页" in main, "第二页没提取到"
assert "第二页备注：客户甲" in main, "第二页正文没提取到"
assert "组合形状内：客户甲" in main, "组合形状文本没提取到"
assert "{{" not in main, "提取产物里还有占位符"
other = open(next(p for p in files if "附件" in p), encoding="utf-8").read()
assert "客户甲 的附件" in other, "第二份文件内容不对:\n" + other[:200]
PY
check "两个 .txt 含页头/正文/表格/组合形状文本" python3 "$TMP/check_extract.py" "$OUT/文本" "$TMP/extract.json"

say "正例 · replace-image 换图（数量不变、内容变、版面不变）"
node bin/cli.mjs replace-image --input "$OUT/fill" --out "$OUT/换图" \
  --shape Logo --image "$FIX/new-logo.png" --json >"$TMP/image.json" 2>&1
RC=$?
[ "$RC" -eq 0 ] && ok "replace-image 退出 0" || bad "replace-image 退出 $RC"
cat > "$TMP/check_image.py" <<'PY'
import glob, hashlib, json, os, sys
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

out, json_path, orig_sha = sys.argv[1], sys.argv[2], sys.argv[3]
data = json.load(open(json_path, encoding="utf-8"))
assert data.get("ok") is True, data
assert data["replaced"] == 1, "期望替换 1 处，实际 %s" % data["replaced"]
assert data["skipped"] == 1, "没有该形状的文件应跳过 1 个: %s" % data["skipped"]
entry = {os.path.basename(f["input"]): f for f in data["files"]}
assert entry["附件.pptx"]["replaced"] == 0 and entry["附件.pptx"].get("skipped"), entry["附件.pptx"]


def pics(prs):
    for slide in prs.slides:
        for sh in slide.shapes:
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                yield sh


path = os.path.join(out, "报价单.pptx")
prs = Presentation(path)
found = list(pics(prs))
assert len(found) == 1, "图片数量应该保持 1，实际 %d" % len(found)
pic = found[0]
sha = hashlib.sha256(pic.image.blob).hexdigest()
assert sha != orig_sha, "图片内容没变（sha 相同）"
assert pic.name == "Logo", "形状名被改了: %s" % pic.name
assert (pic.width, pic.height) == (1371600, 1371600), "图片尺寸/版面被改了: %s x %s" % (pic.width, pic.height)
assert pic.image.content_type == "image/png", pic.image.content_type
texts = "\n".join(sh.text_frame.text for slide in prs.slides for sh in slide.shapes if sh.has_text_frame)
assert "客户甲" in texts, "换图不应影响文字"
# 另一份没有图片的文件也要照样输出
assert os.path.exists(os.path.join(out, "附件.pptx")), "第二份文件没有输出"
PY
check "图片数量不变、哈希已变、形状名与尺寸保留" \
  python3 "$TMP/check_image.py" "$OUT/换图" "$TMP/image.json" "$ORIG_IMG_SHA"

say "正例 · info 页数/形状数/图片数"
node bin/cli.mjs info --input "$FIX" --json >"$TMP/info.json" 2>&1
RC=$?
[ "$RC" -eq 0 ] && ok "info 退出 0" || bad "info 退出 $RC"
cat > "$TMP/check_info.py" <<'PY'
import json, os, sys
data = json.load(open(sys.argv[1], encoding="utf-8"))
assert data.get("ok") is True, data
by = {os.path.basename(f["input"]): f for f in data["files"]}
assert by["报价单.pptx"]["pages"] == 2, by["报价单.pptx"]
assert by["报价单.pptx"]["pictures"] == 1, by["报价单.pptx"]
assert by["报价单.pptx"]["shapes"] >= 5, by["报价单.pptx"]
assert by["附件.pptx"]["pages"] == 1, by["附件.pptx"]
assert by["附件.pptx"]["pictures"] == 0, by["附件.pptx"]
assert data["pages"] == 3, data["pages"]
PY
check "页数/形状数/图片数与实际一致" python3 "$TMP/check_info.py" "$TMP/info.json"
node bin/cli.mjs info --input "$FIX/附件.pptx" >"$TMP/info-human.txt" 2>&1
grep -q "1 页" "$TMP/info-human.txt" && ok "info 人类可读输出含页数（支持单文件）" || bad "info 人类可读输出不对"

say "正例 · 不覆盖同名产物 & 默认输出目录"
node bin/cli.mjs fill --input "$FIX" --out "$OUT/fill" --set client=客户乙 >/dev/null 2>&1
[ -f "$OUT/fill/报价单.pptx" ] && [ -f "$OUT/fill/报价单_1.pptx" ] \
  && ok "同名产物自动加序号（报价单_1.pptx）" || bad "同名产物被覆盖或未加序号"
mkdir -p "$TMP/默认目录" && cp "$FIX/附件.pptx" "$TMP/默认目录/"
node bin/cli.mjs extract-text --input "$TMP/默认目录" >/dev/null 2>&1
[ -f "$TMP/默认目录/_输出/附件.txt" ] && ok "未给 --out 时默认写到 <输入目录>/_输出" || bad "默认输出目录不对"
node bin/cli.mjs fill --input "$FIX/附件.pptx" --out "$TMP/单文件" --set client=客户丙 >/dev/null 2>&1
[ -f "$TMP/单文件/附件.pptx" ] && ok "支持单个 .pptx 输入" || bad "单文件输入失败"

say "正例 · --demo 自造样例跑完整流程"
node bin/cli.mjs --demo --json >"$TMP/demo.json" 2>&1
RC=$?
[ "$RC" -eq 0 ] && ok "--demo 退出 0" || bad "--demo 退出 $RC"
cat > "$TMP/check_demo.py" <<'PY'
import json, os, sys
data = json.load(open(sys.argv[1], encoding="utf-8"))
assert data.get("ok") is True and data.get("action") == "demo", data
steps = data.get("steps", [])
assert len(steps) == 4, "演示步骤不足 4 个: %s" % [s.get("step") for s in steps]
arts = data.get("artifacts", [])
assert len(arts) >= 6, "演示产物不足: %d" % len(arts)
for path in arts:
    assert os.path.exists(path), "演示产物不存在: " + path
print(data["demoRoot"])
PY
DROOT=$(python3 "$TMP/check_demo.py" "$TMP/demo.json" 2>"$TMP/last_err")
if [ $? -eq 0 ]; then ok "fill/replace-image/extract-text/info 四步跑通，产物均存在"; else bad "demo 产物校验失败"; fi
[ -n "${DROOT:-}" ] && rm -rf "$DROOT"

say "反例 · 坏输入必须非 0 退出并说明缺什么"
NEG=$(node bin/cli.mjs extract-text --input "$TMP/根本不存在" 2>&1); RC=$?
[ "$RC" -ne 0 ] && ok "不存在的输入 → 非 0 退出（$RC）" || bad "不存在的输入仍退出 0"
echo "$NEG" | grep -q "不存在" && ok "提示「输入不存在」" || bad "缺少错误提示：$NEG"
mkdir -p "$TMP/空目录"
NEG=$(node bin/cli.mjs info --input "$TMP/空目录" 2>&1); RC=$?
[ "$RC" -ne 0 ] && ok "空目录 → 非 0 退出（$RC）" || bad "空目录仍退出 0"
echo "$NEG" | grep -q "没有 .pptx" && ok "提示「目录里没有 .pptx」" || bad "空目录提示不对：$NEG"
NEG=$(node bin/cli.mjs fill --input "$FIX" 2>&1); RC=$?
[ "$RC" -ne 0 ] && ok "缺 --set → 非 0 退出（$RC）" || bad "缺 --set 仍成功"
echo "$NEG" | grep -q "缺少 --set" && ok "提示「缺少 --set」" || bad "缺 --set 提示不对：$NEG"
NEG=$(node bin/cli.mjs fill --input "$FIX" --set 客户甲 2>&1); RC=$?
[ "$RC" -ne 0 ] && ok "--set 不是 key=value → 非 0 退出（$RC）" || bad "--set 格式错误仍成功"
echo "$NEG" | grep -q "key=value" && ok "提示 key=value 形式" || bad "格式提示不对：$NEG"
NEG=$(node bin/cli.mjs replace-image --input "$FIX" --shape Logo 2>&1); RC=$?
[ "$RC" -ne 0 ] && ok "replace-image 缺 --image → 非 0 退出（$RC）" || bad "缺 --image 仍成功"
echo "$NEG" | grep -q "缺少 --image" && ok "提示「缺少 --image」" || bad "缺 --image 提示不对：$NEG"
NEG=$(node bin/cli.mjs replace-image --input "$FIX" --shape 不存在的形状 --image "$FIX/new-logo.png" 2>&1); RC=$?
[ "$RC" -ne 0 ] && ok "形状名不匹配 → 非 0 退出（$RC）" || bad "形状名不匹配仍成功"
echo "$NEG" | grep -q "没有名为" && ok "提示「没有名为 X 的图片形状」" || bad "形状名提示不对：$NEG"
NEG=$(node bin/cli.mjs replace-image --input "$FIX" --shape 说明 --image "$FIX/new-logo.png" 2>&1); RC=$?
[ "$RC" -ne 0 ] && ok "形状不是图片 → 非 0 退出（$RC）" || bad "非图片形状仍被替换"
echo "$NEG" | grep -q "不是图片" && ok "提示「不是图片」" || bad "非图片提示不对：$NEG"
NEG=$(node bin/cli.mjs nosuch --input "$FIX" 2>&1); RC=$?
[ "$RC" -ne 0 ] && ok "未知子命令 → 非 0 退出（$RC）" || bad "未知子命令仍退出 0"
echo "$NEG" | grep -q "未知子命令" && ok "提示「未知子命令」" || bad "未知子命令提示不对"
NEG=$(node bin/cli.mjs extract-text 2>&1); RC=$?
[ "$RC" -ne 0 ] && ok "缺 --input → 非 0 退出（$RC）" || bad "缺 --input 仍退出 0"
NEG=$(node bin/cli.mjs info --input "$FIX/old-logo.png" 2>&1); RC=$?
[ "$RC" -ne 0 ] && ok "非 .pptx 文件 → 非 0 退出（$RC）" || bad "非 .pptx 文件仍成功"
node bin/cli.mjs extract-text --input "$TMP/根本不存在" --json >"$TMP/err.json" 2>&1; RC=$?
[ "$RC" -ne 0 ] && ok "--json 模式下错误也非 0 退出" || bad "--json 错误仍退出 0"
python3 -c "import json,sys; d=json.load(open(sys.argv[1])); assert d['ok'] is False and d['message']" "$TMP/err.json" \
  >/dev/null 2>"$TMP/last_err" && ok "--json 错误输出可解析且含 message" || bad "--json 错误输出不可解析"

AFTER2=$(shasum -a 256 "$FIX"/*.pptx | shasum -a 256)
[ "$BEFORE" = "$AFTER2" ] && ok "全流程结束后输入文件仍未改动（sha256 一致）" || bad "全流程后输入文件被改动"

echo ""
echo "  ────────────────────────────────────────────────────────"
echo "  通过 $PASS · 失败 $FAIL"
echo ""
[ "$FAIL" -eq 0 ] || exit 1
