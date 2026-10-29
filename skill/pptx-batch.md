# 深鲸 PPT 批量（dsh-pptx-batch）

**什么时候用**：批量处理 PPT。用户会说「PPT 文字提出来」「导出里面的图」「一共多少页」。

```bash
<cli> --op extract-text --input <目录或文件>
<cli> --op extract-images --input <目录> [--out 图片目录]
<cli> --op list --input <目录>
```

需要 python3 + python-pptx；只读不覆盖。
