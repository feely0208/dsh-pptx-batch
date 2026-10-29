import { spawnSync } from 'node:child_process';
const PY = String.raw`
import sys, json, os
op = sys.argv[1] if len(sys.argv) > 1 else 'extract-text'
inp = sys.argv[2] if len(sys.argv) > 2 else ''
out = sys.argv[3] if len(sys.argv) > 3 else ''
try:
    from pptx import Presentation
except Exception:
    print(json.dumps({"ok": False, "text": "需要 python3 + python-pptx：pip3 install python-pptx"}, ensure_ascii=False)); sys.exit(0)

def fail(msg):
    print(json.dumps({"ok": False, "text": msg}, ensure_ascii=False)); sys.exit(0)
files = []
if os.path.isdir(inp):
    for n in sorted(os.listdir(inp)):
        if n.lower().endswith('.pptx') and not n.startswith('~$'): files.append(os.path.join(inp, n))
elif inp.lower().endswith('.pptx'):
    files = [inp]
if not files:
    print(json.dumps({"ok": False, "text": "没有找到 .pptx 文件"}, ensure_ascii=False)); sys.exit(0)
lines = []
try:
  if op == 'list':
    for f in files:
        lines.append(os.path.basename(f) + ': ' + str(len(Presentation(f).slides)) + ' 页')
  elif op == 'extract-images':
    if not out: out = os.path.join(os.path.dirname(files[0]), '_图片')
    os.makedirs(out, exist_ok=True); n = 0
    for f in files:
        for si, s in enumerate(Presentation(f).slides, 1):
            for sh in s.shapes:
                if sh.shape_type == 13:
                    n += 1
                    open(os.path.join(out, str(si) + '-' + str(n) + '.png'), 'wb').write(sh.image.blob)
    lines.append('导出 ' + str(n) + ' 张图片到 ' + out)
  else:
    for f in files:
        prs = Presentation(f); lines.append('## ' + os.path.basename(f))
        for i, s in enumerate(prs.slides, 1):
            texts = [sh.text_frame.text.strip() for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
            lines.append('### 第 ' + str(i) + ' 页')
            lines.extend(texts if texts else ['（无文字）'])
except Exception as e:
    fail('无法读取 PPT（可能不是有效的 .pptx）：' + str(e)[:200])
print(json.dumps({"ok": True, "text": chr(10).join(lines)}, ensure_ascii=False))
`;
export function run({ op = 'extract-text', input = '', out = '' } = {}) {
  if (!input) return { ok: false, text: '请填写 PPT 文件或目录' };
  const r = spawnSync('python3', ['-c', PY, op, input, out], { encoding: 'utf8' });
  try { return JSON.parse((r.stdout || '').trim() || '{}'); } catch { return { ok: false, text: String((r.stderr || r.stdout || '执行失败')).slice(-400) }; }
}
