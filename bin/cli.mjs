#!/usr/bin/env node
import { run } from '../lib/core.mjs';

const args = process.argv.slice(2);
const val = (k, d = '') => { const i = args.indexOf('--' + k); return i >= 0 && args[i + 1] && !args[i + 1].startsWith('--') ? args[i + 1] : d; };
if (!args.length || args.includes('--help')) {
  console.log(`dsh-pptx-batch —— 深鲸 PPT 批量

用法：dsh-pptx-batch --op <extract-text,extract-images,list> [参数]
参数：input,out
示例：dsh-pptx-batch --op extract-text --input ./提案`);
  process.exit(0);
}
const input = { op: val('op', 'extract-text') };
for (const k of ["input", "out"]) input[k] = val(k);
const r = await run(input);
if (!r.ok) { console.error('❌', r.text || '失败'); process.exit(1); }
console.log(r.text || '');
process.exit(0);
