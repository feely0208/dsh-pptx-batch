import { createPlugin, pkgRootFrom } from '@deepwhale-cn/dsh-guard-core/plugin';
import { registerApi } from './api.js';

const base = createPlugin({
  name: 'pptx-batch',
  pkgRoot: pkgRootFrom(import.meta.url),
  skillFile: 'skill/pptx-batch.md',
  binFile: 'bin/cli.mjs',
  description: '深鲸 PPT 批量：提取幻灯片文字、导出图片、统计页数，一次处理整目录',
  whenToUse: '当用户要批量处理 PPT 时使用：提取文字、导出图片、统计页数。典型说法：「把这批 PPT 的文字都提出来」「PPT 里的图导出一下」。',
});
export const name = base.name;
export const inject = [...base.inject, 'webServer'];
export function apply(ctx) {
  base.apply(ctx);
  try { if (ctx.webServer?.register) registerApi(ctx); } catch (e) { ctx.logger?.warn?.('[pptx-batch] 面板出口注册失败: ' + (e?.message || e)); }
}
