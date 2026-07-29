# 大家的日语长期词库工程

以 `data/vocabulary.json` 为唯一数据源，生成阅读词库、随机默写 Excel、词库 PDF、默写 PDF 和 Word。数据与程序分离，修改 JSON 后无需重新编译程序。

## 环境

- Python 3.12+
- Windows、macOS

```bash
python -m pip install -r requirements.txt
```

## 常用命令

```bash
# 重新生成词库和默写 Excel
python main.py excel

# 导出词库 PDF / Word
python main.py pdf
python main.py word

# 从 Lesson 1、3、5 抽 50 道日译中，末尾附答案
python main.py exam-pdf --lessons 1,3,5 --count 50 --answers

# 全部课程、全部可默写词条、中译日
python main.py exam-pdf --count all --direction zh_to_jp
```

默写 PDF 直接按实际题数从 JSON 生成，不打印 Excel 的预留空行，因此不会产生大量空白页。生成文件位于 `output/`。

## JSON 字段

`id`、`lesson`、`order`、`type`、`japanese`、`kana`、`kanji`、`chinese`、`remark`。

`type` 支持：

- `word`：单词
- `expression`：固定表达
- `example`：例句，不进入默认默写题库

追加课程时，按相同字段添加 Lesson26、Lesson27 等记录，再运行生成命令即可。

## GitHub 建议

- 提交源码和 `data/vocabulary.json`。
- 教材 PDF、EXE、ZIP、OCR 临时文件及生成产物不进入 Git。
- Windows 免安装版和跨平台压缩包建议放到 GitHub Releases。
- 请确认公开发布词库数据和教材内容符合版权要求。
