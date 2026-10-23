# 日语词库与随机默写

**本项目所有内容均由 AI 生成，包括代码、词库整理、文档及发布说明。词条可能存在错误，请结合教材或词典核对。**

内置《大家的日语》第1～50课，支持选择课程、双向默写、随机抽题，以及导出 Excel 和 Word。可扩充课程或使用自建词库。

## 下载与使用

从 [最新版本](https://github.com/ganseed/minna-no-nihongo-vocabulary/releases/latest) 下载 Windows 64位或 Mac（Apple Silicon／M系列）压缩包，完整解压后打开“大家的日语词库工具”。Windows 请保留包内目录结构。Mac 应用内置词库和宏模板，也可单独移动到“应用程序”；选择其他 JSON 可使用自己的词库。

1. 选择词库、课程、默写方向和题数。
2. 课程支持单课 `3`、多课 `1,3,5`、范围 `1-5` 或 `全部`；题数支持正整数或 `全部`。
3. 生成 Excel 或 Word。Mac 应用默认保存到 `~/Documents/日语词库/output`，Windows 和源码运行默认保存在项目的 `output` 文件夹；也可自行选择位置。

每次生成重新随机，同一份试卷内不重复抽题；题数不足时按实际数量生成。答案在作答栏右侧，少题无多余空框，多题保留边框。同名输出会被覆盖，生成前请关闭已打开的文件。

## Excel 宏版

使用支持 VBA 的 Microsoft Excel 打开 `默写-宏版.xlsm` 并启用宏，在“设置”页修改课程、方向和题数，点击“随机生成”。“默写”页也有重新随机按钮，每次点击都会清空旧作答。

宏版自带生成时所选 JSON 中的完整可用题库，即使初次只选一课，之后也能切换其他课。**宏版内多课请用 `1,2,3`，不支持 `1-3` 范围写法。**

## 扩充词库

修改 `data/vocabulary.json` 或选择其他 JSON，再重新生成文件即可。课程列表从词库读取，不限制为50课；已有宏版不会自动同步新增词条。

兼容原有 JSON 数组，也支持自定义名称和分组：

```json
{
  "title": "旅行日语",
  "unit_label": "主题",
  "records": [
    {"id": "travel-001", "lesson": 1, "japanese": "駅", "chinese": "车站"}
  ]
}
```

词条必填 `id`（唯一）、`japanese`、`chinese`。`lesson` 为正整数，省略时归入分组1；其他扩展字段会保留在导出的 `Vocabulary` 数据页。仅中文释义非空的词条参与默写。

## 源码运行

使用 Python 3.12 或更新版本：

```sh
python -m pip install -r requirements.txt
python desktop_app.py
```

测试：`python -m unittest discover -s tests -v`。本地打包使用 `build_windows_app.ps1` 或 `build_macos_app.sh`；推送 `v*` 标签后自动测试并发布双平台 ZIP。
