#!/usr/bin/env node
/**
 * PPT 批量处理 CLI
 *
 * 形态：
 *   dsh-pptx-batch --help                       （无参数也打印 help，退出 0）
 *   dsh-pptx-batch --demo                       （自造样例跑完整流程，退出 0）
 *   dsh-pptx-batch <action> --input <文件或目录> [--out <目录>] [--json]
 *
 * 真实逻辑全部委托 python3 + python-pptx（lib/pptx_tool.py），
 * Node 只负责参数校验、调用与人类可读输出。
 */
import { spawnSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const PKG_ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const PY_SCRIPT = join(PKG_ROOT, 'lib', 'pptx_tool.py');
const TITLE = 'PPT 批量处理';
const CMD = 'dsh-pptx-batch';
const ACTIONS = {
  fill: '按 --set key=value 批量填模板里的 {{key}} 占位符',
  'replace-image': '把名为 --shape 的图片形状换成 --image 指定的图片',
  'extract-text': '提取每页文本（标题 + 正文 + 表格）到同名 .txt',
  info: '输出每份 PPT 的页数 / 形状数 / 图片数',
};
const PY_HINT = [
  '需要 python3 + python-pptx。安装：',
  '  python3 -m pip install python-pptx',
  '没有 python3 时：',
  '  macOS: brew install python3',
  '  Ubuntu: sudo apt install python3',
].join('\n');
const LINE = '─'.repeat(56);

function help() {
  console.log(`${TITLE}

用法：
  ${CMD} <action> --input <文件或目录> [--out <目录>] [--json]
  ${CMD} --demo
  ${CMD} --help

子命令：
${Object.entries(ACTIONS).map(([k, v]) => `  ${k.padEnd(14)}${v}`).join('\n')}

参数：
  --input, -i   .pptx 文件或目录（目录递归处理）
  --out, -o     输出目录；默认 <输入目录>/_输出；同名文件不覆盖，自动加序号
  --set         key=value，可重复；对应幻灯片里的 {{key}} 占位符
  --shape       要替换的图片形状名（精确匹配，大小写敏感）
  --image       新图片文件（png / jpg / gif / bmp / tiff）
  --json        输出 JSON（给脚本 / CI 用）
  --demo        自造样例 PPT，跑完整流程并打印产物路径

示例：
  ${CMD} fill --input ./模板 --set 客户=客户甲 --set 金额=128000
  ${CMD} replace-image --input ./已填 --shape Logo --image ./logo-新.png
  ${CMD} extract-text --input ./模板 --out ./文本
  ${CMD} info --input ./模板 --json

依赖：python3 + python-pptx（缺失时会提示安装命令）`);
  process.exit(0);
}

function fail(message, hint = '', code = 2, json = false) {
  if (json) {
    console.log(JSON.stringify({ ok: false, code, message, hint }, null, 2));
  } else {
    console.error(`❌ ${message}`);
    if (hint) console.error(`\n${hint}`);
  }
  process.exit(code);
}

function findPython() {
  const candidates = [
    process.env.DSH_PYTHON,
    process.env.PPTX_BATCH_PYTHON,
    '/opt/homebrew/bin/python3',
    'python3',
    '/usr/local/bin/python3',
    '/usr/bin/python3',
  ].filter(Boolean);
  for (const cmd of candidates) {
    const r = spawnSync(cmd, ['-c', 'import sys; print(sys.version_info[0])'], {
      encoding: 'utf8',
      timeout: 10000,
    });
    if (!r.error && r.status === 0) return cmd;
  }
  return null;
}

function parseArgs(argv) {
  const opt = {
    action: null, input: null, out: null, set: [], shape: null, image: null,
    json: false, demo: false,
  };
  const VALUE_FLAGS = new Set(['--input', '-i', '--out', '-o', '--set', '--shape', '--image']);
  for (let i = 0; i < argv.length; i++) {
    const raw = argv[i];
    const eq = raw.indexOf('=');
    const flag = eq > 0 ? raw.slice(0, eq) : raw;
    const inline = eq > 0 ? raw.slice(eq + 1) : undefined;
    if (flag === '--json') {
      opt.json = true;
      continue;
    }
    if (flag === '--demo') {
      opt.demo = true;
      continue;
    }
    if (flag === '--help' || flag === '-h') help();
    if (VALUE_FLAGS.has(flag)) {
      let value = inline;
      if (value === undefined) {
        value = argv[i + 1];
        if (value === undefined || value.startsWith('--')) {
          fail(`参数 ${flag} 缺少取值`, `例：${flag} <值>`);
        }
        i += 1;
      }
      if (flag === '--input' || flag === '-i') opt.input = value;
      else if (flag === '--out' || flag === '-o') opt.out = value;
      else if (flag === '--shape') opt.shape = value;
      else if (flag === '--image') opt.image = value;
      else opt.set.push(value);
      continue;
    }
    if (flag.startsWith('-')) {
      fail(`未知参数：${raw}`, `用 ${CMD} --help 查看支持的参数。`);
    }
    if (!opt.action) {
      opt.action = raw;
      continue;
    }
    fail(`多余的参数：${raw}`, '一次只能执行一个子命令。');
  }
  return opt;
}

function human(res) {
  console.log('');
  console.log(`  ${TITLE} · ${res.action}`);
  console.log(`  ${LINE}`);
  if (res.input) console.log(`  输入       ${res.input}`);
  if (res.out) console.log(`  输出目录   ${res.out}`);
  if (res.image) console.log(`  新图片     ${res.image}`);
  console.log('');
  for (const f of res.files || []) {
    switch (res.action) {
      case 'fill': {
        const keys = Object.entries(f.keys || {})
          .map(([k, v]) => `${k}×${v}`).join(' ');
        console.log(`  ✓ ${f.input_name} → ${f.output_name} · 替换 ${f.replacements} 处${keys ? `（${keys}）` : ''}`);
        break;
      }
      case 'replace-image':
        if (f.replaced === 0) {
          console.log(`  – ${f.input_name} → ${f.output_name} · 跳过（没有该形状）`);
        } else {
          console.log(`  ✓ ${f.input_name} → ${f.output_name} · 替换图片 ${f.replaced} 处`);
        }
        break;
      case 'extract-text':
        console.log(`  ✓ ${f.input_name} → ${f.output_name} · ${f.slides} 页 · ${f.chars} 字符`);
        break;
      case 'info':
        console.log(`  ✓ ${f.input_name} · ${f.pages} 页 · ${f.shapes} 个形状 · ${f.pictures} 张图片`);
        break;
      default:
        console.log(`  ✓ ${f.input_name} → ${f.output_name}`);
    }
  }
  if (res.unfilled && res.unfilled.length) {
    console.log(`\n  ⚠ 还有 ${res.unfilled.length} 个占位符没给值：${res.unfilled.join('、')}`);
    console.log('    确认 --set 的 key 和模板里的 {{key}} 完全一致（大小写、空格）。');
  }
  if (res.summary) console.log(`\n  ${res.summary}`);
  if (res.artifacts && res.artifacts.length) {
    console.log('\n  产物：');
    for (const a of res.artifacts) console.log(`    ${a}`);
  }
  console.log('');
}

function humanDemo(res) {
  console.log('');
  console.log(`  ${TITLE} · 演示（--demo）`);
  console.log(`  ${LINE}`);
  console.log(`  演示目录   ${res.demoRoot}`);
  console.log('');
  for (const s of res.steps || []) {
    console.log(`  ✓ ${String(s.step).padEnd(6)} ${s.action.padEnd(14)} ${s.summary}`);
  }
  console.log('\n  产物：');
  for (const a of res.artifacts || []) console.log(`    ${a}`);
  console.log('');
}

function callPython(py, action, payload) {
  const r = spawnSync(py, [PY_SCRIPT, action, JSON.stringify(payload)], {
    encoding: 'utf8',
    maxBuffer: 128 * 1024 * 1024,
  });
  if (r.error) {
    fail(`调用 python 失败：${r.error.message}`, PY_HINT, 3);
  }
  const lines = (r.stdout || '').trim().split('\n').filter(Boolean);
  let result = null;
  try {
    result = JSON.parse(lines[lines.length - 1] || '');
  } catch {
    /* 下面统一报错 */
  }
  if (!result) {
    const detail = (r.stderr || '').trim().split('\n').slice(-5).join('\n');
    fail('Python 处理失败（没有拿到有效结果）', detail || PY_HINT, r.status || 1);
  }
  return result;
}

const argv = process.argv.slice(2);
if (argv.length === 0) help();
if (argv.includes('--help') || argv.includes('-h')) help();

const opt = parseArgs(argv);
const python = findPython();
if (!python) {
  fail('未找到 python3，无法处理 PPT 文件', PY_HINT, 3, opt.json);
}

if (opt.demo) {
  const res = callPython(python, 'demo', {});
  if (!res.ok) fail(res.message, res.hint, res.code || 1, opt.json);
  if (opt.json) console.log(JSON.stringify(res, null, 2));
  else humanDemo(res);
  process.exit(0);
}

if (!opt.action) {
  fail('缺少子命令', `可用子命令：${Object.keys(ACTIONS).join(' / ')}\n用 ${CMD} --help 查看详情。`, 2, opt.json);
}
if (!Object.hasOwn(ACTIONS, opt.action)) {
  fail(`未知子命令：${opt.action}`, `可用子命令：${Object.keys(ACTIONS).join(' / ')}`, 2, opt.json);
}
if (!opt.input) {
  fail('缺少 --input（文件或目录）', `例：${CMD} ${opt.action} --input ./PPT目录`, 2, opt.json);
}
if (opt.action === 'fill') {
  if (opt.set.length === 0) {
    fail('缺少 --set（要填的占位符）',
      `例：${CMD} fill --input ./模板 --set 客户=客户甲 --set 金额=128000`, 2, opt.json);
  }
  const bad = opt.set.find((s) => !s.includes('=') || !s.split('=')[0].trim());
  if (bad !== undefined) {
    fail(`--set 需要 key=value 形式：${bad}`, '例：--set 客户=客户甲', 2, opt.json);
  }
}
if (opt.action === 'replace-image') {
  if (!opt.shape) {
    fail('缺少 --shape（要替换的图片形状名）',
      `例：${CMD} replace-image --input ./模板 --shape Logo --image ./logo.png`, 2, opt.json);
  }
  if (!opt.image) {
    fail('缺少 --image（新图片文件）',
      `例：${CMD} replace-image --input ./模板 --shape Logo --image ./logo.png`, 2, opt.json);
  }
}

const payload = {
  input: opt.input,
  out: opt.out,
  set: opt.set,
  shape: opt.shape,
  image: opt.image,
};
const res = callPython(python, opt.action, payload);
if (!res.ok) fail(res.message, res.hint, res.code || 1, opt.json);

if (opt.json) console.log(JSON.stringify(res, null, 2));
else human(res);
process.exit(0);
