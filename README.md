# 深鲸 PPT 批量（@deepwhale-cn/dsh-pptx-batch）

提取幻灯片文字 / 导出图片 / 统计页数，一次处理整目录。

## 安装

```bash
dsh plugin --profile web add @deepwhale-cn/dsh-pptx-batch
```

## 用法

```bash
dsh-pptx-batch --op extract-text --input ./提案
dsh-pptx-batch --op extract-images --input ./提案 --out ./图片
dsh-pptx-batch --op list --input ./提案
```

面板：侧栏「深鲸 PPT 批量」。需要 `python3 + python-pptx`；只读，不修改原文件。
