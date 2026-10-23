# 大家的日语词库与随机默写

用于《大家的日语》初级Ⅰ、Ⅱ（第1～50课）的阅读、默写和词库维护，支持 Windows 和 macOS。

## 下载和使用

到 [Releases](https://github.com/ganseed/minna-no-nihongo-vocabulary/releases/latest) 下载适合平台的 ZIP，**完整解压**后双击“大家的日语词库工具”。保留应用、`data`、`template` 和 `output` 的目录关系。

1. 启动时自动选择发布包内置词库，也可选择其他兼容 JSON。
2. 填写课程：单课 `3`，多课 `1,3,5`，连续课程 `1-5`，或 `全部`。
3. 选择中译日／日译中、题数和输出文件夹。生成前会显示可用题数和实际生成题数。
4. 点击“生成 Excel 词库和默写表”或“生成所选课程 Word 词库”。

题数可以填正整数或“全部”。如果所选课程只有60题，即使请求200题，也只生成60题，不补空题。

## 生成结果

| 文件 | 用途 |
| --- | --- |
| 大家的日语Ⅰ词库.xlsx | 所选课程的阅读词库，只创建实际有词条的课程页 |
| 大家的日语Ⅰ默写.xlsx | 本次试卷快照，严格按实际题数生成，答案在独立工作表 |
| 大家的日语Ⅰ默写-宏版.xlsm | 在 Microsoft Excel 中更改设置并反复随机出题 |
| 大家的日语Ⅰ词库.docx | 所选课程的可编辑 Word 词库，单独生成 |

普通版换课程、换方向或重新随机时，请回到桌面工具重新生成。它不需要启用宏。

已有同名文件会在成功生成后替换；如果文件正被 Excel 或 Word 打开，请先关闭再生成。

## 宏版默写

打开宏版，在 Microsoft Excel 提示时按需启用宏。在“设置”工作表中填写：

| 单元格 | 内容 |
| --- | --- |
| B2 | 中译日／日译中 |
| B3 | 单课 `3`，多课 `1,3,5` 或 `全部` |
| B4 | 20、50、100、200 或 全部 |
| B8 | 实际题数，自动更新 |

点击“随机生成”，再查看“默写”页。每次点击会重新洗牌，并清空上一份试卷的作答。

**宏版工作簿内的课程输入不支持 `1-5`，请写成 `1,2,3,4,5`。** 桌面工具支持范围，并会展开成宏版能识别的课次。

- 少题时不显示多余空框，多题时自动补齐边框。表格显示依赖条件格式，请勿删除条件格式规则。
- “默写”页用于作答，“答案”页用于核对。默写页的D列答案默认隐藏。
- 两页分别设置动态打印范围，按实际题数打印。请打印当前“默写”页或“答案”页。
- 宏版保留全词库题目，方便继续切换其他课程；后台“题库”页默认隐藏。

宏功能需要支持 VBA 的 Microsoft Excel。WPS 的宏支持因版本而异。

## 词库维护

统一数据源为 `data/vocabulary.json`。修改后重新生成即可，不需要重新编译应用。支持 UTF-8 和带 BOM 的 UTF-8。

必要字段：`id`（唯一编号）、`lesson`（正整数课次）、`order`（正整数顺序）、`type`、`japanese`、`chinese`。

可选字段：`kana`、`kanji`、`remark`、`part_of_speech`、`pitch_accent`、`dictionary_form`。生成的 `Vocabulary` 工作表保留这些字段。

全部有非空中文释义的记录均可抽题。抽题不重复；“全部”会将所选课程的可用题目全部打乱。

声调沿用现有词库标法，例如 `24` 表示②、④两种，`01` 表示⓪、①两种。动词的声调多数按辞书形记录，不能直接套用到ます形。本词库为自用学习资料，未核验的声调仍应结合词典确认。

## 从源码运行

需要 Python 3.12 或更高版本：

```sh
python -m pip install -r requirements.txt
python desktop_app.py
```

命令行示例：

```sh
python main.py excel --lessons "3" --count 20
python main.py excel --lessons "1,3,5" --count 全部 --direction 日译中
python main.py all --lessons "1-5" --output "./output"
python main.py word --lessons "8"
python main.py excel --data "./data/vocabulary.json" --seed 42
```

macOS 可将 `python` 替换为 `python3`。`--seed` 只用于需要复现同一份试卷时，默认每次随机。

## 开发和发布

运行测试：

```sh
python -m unittest discover -s tests -v
```

测试覆盖实际文件生成、20→200→20题、题数不足、课程筛选、宏版条件边框、动态打印范围、VBA项目保留、Word首屏和词库校验。

本机构建：

- Windows：`./build_windows_app.ps1`
- macOS：`./build_macos_app.sh`

GitHub Actions 提供独立的 Windows、macOS 构建，以及双平台测试。推送 `v*` 标签后，发布工作流先测试，再构建两个平台，并运行打包后生成检查；两个包都成功后，才创建包含双平台 ZIP 的 GitHub Release。发布说明取自 `RELEASE_NOTES.md`。

打包检查验证应用能独立读取外置词库和宏模板、生成默认20题试卷。自动构建不启动 Microsoft Excel 实际执行 VBA。

项目不包含 PDF 导出。生成结果与构建目录默认不提交，发布包通过 GitHub Releases 分发。
