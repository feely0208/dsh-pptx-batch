import { createPlugin, pkgRootFrom } from '@deepwhale-cn/dsh-guard-core/plugin';

export const { name, inject, apply } = createPlugin({
  name: 'pptx-batch',
  pkgRoot: pkgRootFrom(import.meta.url),
  skillFile: 'skill/pptx-batch.md',
  binFile: 'bin/cli.mjs',
  description: 'PPT 批量填模板 / 替换文字与图片 / 导出内容清单（python-pptx）',
  whenToUse: '当用户要批量处理整目录的 PowerPoint（.pptx）时使用：按 {{key}} 批量填模板（客户名、金额、日期）、按形状名批量换图片（Logo）、提取每页文本到 txt、统计页数/形状数/图片数。典型说法：「这套模板有 30 份要按客户发」「把每份 PPT 里的旧 Logo 换掉」「把这批 pptx 的文字导出来核对」。不用于 .ppt 老格式、加密文档、需要保持复杂动画与版式的场景。',
});
