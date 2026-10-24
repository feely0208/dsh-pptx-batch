# @deepwhale-cn/dsh-pptx-batch

[![Listed on dsh-plugin.org](https://dsh-plugin.org/badges/listed.svg)](https://dsh-plugin.org/plugins/feely0208/dsh-pptx-batch)

**整目录的 PPT 一次处理完**：批量填模板占位符、按形状名换图片、提取每页文本、看页数/形状数/图片数 —— 不打开 PowerPoint，不动原文件。

**关键词**：PPT 批量填模板 · pptx 批量替换 · PPT 批量换 Logo · 提取 PPT 文本 · PPT 页数统计 · 汇报批量生成 · 模板占位符替换 · 命令行处理 PPT

## 不装会怎样

30 份给不同客户的报价 PPT，模板里写着 `{{客户}}`、`{{金额}}`：不装就得一份份打开、双击文本框、输入、保存、关掉，
一小时起步，改到第 20 份时手一滑把上一位客户的名字留在下一份里，也没人发现。

五十份材料的 Logo 要换新版：人工做法是逐份删图、插图、再对齐一遍位置，版面还容易歪。

要核对一批 PPT 每页写了什么、各有多少页多少张图：只能一份份翻，翻完还记不住。

这些都属于「机械但量大、还容易漏」的活。更麻烦的是直接让 AI 改原文件，改坏了没有退路 ——
这个插件**只读输入**，产物写到单独的输出目录，同名文件自动加序号，不覆盖任何东西。

## 能做什么

| 子命令 | 作用 | 产物（可核对） |
|---|---|---|
| `fill` | 按 `--set key=value` 批量填幻灯片里的 `{{key}}` 占位符（文本框、占位符、表格单元格、组合形状） | 填好的 `.pptx` 副本；JSON 里给出每份替换了几处、哪个 key 几处 |
| `replace-image` | 把名为 `--shape` 的图片形状换成 `--image` 指定的图片 | 换图后的 `.pptx` 副本；图片数量不变、位置尺寸不变、形状名不变 |
| `extract-text` | 提取每页文本：标题 + 正文 + 表格 + 备注 | 每份 PPT 一个同名 `.txt`，按「第 N 页」分块 |
| `info` | 输出每份 PPT 的页数、形状数、图片数 | 终端表格；`--json` 给脚本用 |

共同约定：

- **不修改输入文件**，默认输出到 `<输入目录>/_输出`，可用 `--out` 指定；
- **同名文件不覆盖**：已存在时自动变成 `简报_1.pptx`、`简报_2.pptx`；
- 目录递归处理，自动跳过 PowerPoint 的 `~$` 临时文件和 `_输出` 目录；
- `--json` 输出结构化结果，方便脚本 / CI 断言；出错时退出码非 0 并说明缺什么；
- `fill` 填完会报出**还剩哪些 `{{...}}` 没给值**，模板里拼错 key 一眼能看出来。

```bash
# 批量填模板：一份模板 + 一个客户一套值
dsh-pptx-batch fill --input ./模板 --out ./已填 \
  --set 客户=客户甲 --set 金额=128000 --set 日期=2026-10-24

# 批量换 Logo（形状名要和 PowerPoint「选择窗格」里显示的一致）
dsh-pptx-batch replace-image --input ./已填 --out ./换图 \
  --shape Logo --image ./logo-新.png

# 提取每页文本
dsh-pptx-batch extract-text --input ./模板 --out ./文本

# 看一眼结构
dsh-pptx-batch info --input ./模板 --json
```

`fill` 保留格式的方式：占位符整段落在一个 run 里时，这个 run 的字体、字号、颜色完全不动；
只有跨多个 run 的占位符（同一段字被 PowerPoint 拆成几块）才会把这几段并到起始 run 的格式上。

## 不能做什么

- ❌ 只支持 `.pptx`；`.ppt` 老格式、`.pptm`、加密/需要密码打开的文档不支持 —— `.ppt` 请先用 PowerPoint/WPS 另存为 `.pptx`。
- ❌ `fill` **不处理**图表文字、SmartArt 文字、图片里的文字、批注、母版/版式上的文字，也不跨段落匹配。
- ❌ `fill` 是**字面匹配**（`{{key}}`，key 大小写敏感，允许 `{{ key }}` 这种空格），不支持正则、不支持一次给同名 key 多个值。
- ❌ **不保证排版 100% 不变**：跨 run 的占位符替换会合并该段文字的 run，段落内局部加粗/字号/颜色可能丢失；
  重要模板请先跑一份对比再批量用。
- ❌ `replace-image` 只换图片内容：新图比例和原图不同时会被**拉伸**（裁剪设置沿用原来的）；
  只按形状名精确匹配，不支持按内容或页码挑选，同名形状会全部替换。
- ❌ `extract-text` 不提取图表数据、SmartArt、页码与母版文字；图片只留位置不留文字（没有 OCR）。
- ❌ `info` 的「形状数」是每页顶层形状数之和（组合形状算 1 个，其内部形状不计入）。
- ❌ 不是格式转换器：不导出 PDF、不做 OCR、不做批量重命名。
- ✅ 全部在本机由 python-pptx 处理，不联网、不上传文件。

## 安装

```bash
dsh plugin --profile web add @deepwhale-cn/dsh-pptx-batch
```

运行依赖本机的 **python3 + python-pptx**（插件不会替你装）。缺依赖时 CLI 会直接打印安装命令：

```bash
python3 -m pip install python-pptx
# macOS 没有 python3 时
brew install python3
```

## 用法

```bash
dsh-pptx-batch --help
dsh-pptx-batch --demo        # 自造样例 PPT 跑完整流程，打印产物路径
dsh-pptx-batch <action> --input <文件或目录> [--out <目录>] [--json]
```

`--demo` 会在系统临时目录里造两份带占位符、表格和图片形状的样例 PPT，
依次跑 fill / replace-image / extract-text / info，不动你的任何文件，跑完直接看产物路径就知道它能干什么。

退出码：`0` 成功 · `1` 运行失败 · `2` 参数或输入问题 · `3` 缺依赖。

## License

MIT © 温州深鲸智能科技有限公司
