# 大家的日语长期词库工程

本项目以 `data/vocabulary.json` 为统一数据源，用于维护《大家的日语 初级 I》词库，并生成阅读词库、随机默写工作簿和 Word 文档。

当前词库包含 Lesson 01～50。后续可以继续追加课程，不需要重做工程。

## 主要功能

- 按课生成词库 Excel
- 默认使用“中译日”默写
- 支持单课、多课或全部课程
- 支持 20、50、100、200 题或所选课程全部题目
- 题目范围包含所选课程中所有带中文内容的记录
- 宏版 Excel 可在工作簿内点击“随机生成”重新洗牌
- 导出可编辑的 Word 词库
- 桌面工具可选择内置或其他兼容的 JSON 文件
- Word 与 Excel 使用独立生成按钮
- 支持 Windows 和 macOS
- 不包含 PDF 导出功能

## 直接使用宏版默写

打开：

```text
output/大家的日语Ⅰ默写-宏版.xlsm
```

首次打开时，请在 Microsoft Excel 中点击“启用内容”或“启用宏”。

在“设置”工作表中设置：

| 单元格 | 功能 | 填写示例 |
| --- | --- | --- |
| B2 | 默写方向 | `中译日` 或 `日译中` |
| B3 | 课程范围 | 单课 `3`；多课 `1,3,5`；全部课程 `全部` |
| B4 | 题目数量 | `20`、`50`、`100`、`200` 或 `全部` |

默认设置为“中译日、全部课程、20 题”。

设置完成后点击“随机生成”。每点击一次都会重新打乱题目，不使用固定随机种子。生成结果位于“默写”工作表，包含题目、作答栏和答案。

### 宏版兼容性

- Windows Excel：支持
- Mac Excel：支持，首次打开时需要允许宏
- Microsoft 365：支持
- WPS：宏支持因版本而异，不保证按钮可用

如果 Windows 提示文件来自网络并阻止宏，请关闭工作簿，在文件上右键选择“属性”，勾选“解除锁定”，然后重新打开并启用宏。

## 从 GitHub 下载

可以使用 Git：

```bash
git clone <你的 GitHub 仓库地址>
cd minna-no-nihongo-vocabulary
```

也可以在 GitHub 页面选择 `Code` → `Download ZIP`，解压后进入项目目录。

如果只需要默写，可以直接使用仓库中的宏版工作簿，不需要安装 Python。

## 桌面工具

Windows 和 macOS 发布包使用相同界面，提供：

- “选择 JSON 文件”：默认指向发布包内的 `data/vocabulary.json`，也可以使用其他兼容词库
- “生成 Excel 词库 + 默写模板”：同时生成阅读词库和普通默写工作簿
- “生成 Word 词库”：单独生成可编辑 Word 文档
- “打开输出文件夹”：查看生成结果

两个平台的发布包结构一致：

```text
大家的日语词库-Windows/ 或 大家的日语词库-macOS/
├── 大家的日语词库工具.exe 或 大家的日语词库工具.app
├── data/
│   └── vocabulary.json
├── output/
└── README.md
```

应用、`data` 和 `output` 应保持在同一个发布目录中。macOS 包已经自带默认 `data/vocabulary.json`，不需要手动复制。

## 使用源码生成文件

需要 Python 3.12 或更高版本。

Windows：

```powershell
python -m pip install -r requirements.txt
python main.py excel
```

macOS：

```bash
python3 -m pip install -r requirements.txt
python3 main.py excel
```

也可以指定其他 JSON 和输出目录：

```powershell
python main.py excel --data "D:/词库/其他词库.json" --output "D:/词库/生成结果"
python main.py word --data "D:/词库/其他词库.json" --output "D:/词库/生成结果"
```

生成结果位于 `output/`：

- `大家的日语Ⅰ词库.xlsx`
- `大家的日语Ⅰ默写.xlsx`

普通 `.xlsx` 不含宏。需要新的随机顺序时，可以再次执行 `excel` 命令；如果使用宏版，直接点击工作簿内的“随机生成”即可。

## 导出 Word

Windows：

```powershell
python main.py word
```

macOS：

```bash
python3 main.py word
```

生成结果：

```text
output/大家的日语Ⅰ词库.docx
```

同时生成 Excel 和 Word：

```powershell
python main.py all
```

## 修改词库

长期维护时只需编辑：

```text
data/vocabulary.json
```

建议使用 VS Code，并保持 UTF-8 编码。每条记录包含：

- `id`：唯一编号，例如 `L01-001`
- `lesson`：课次
- `order`：本课顺序
- `type`：`word`、`expression` 或 `example`
- `japanese`：日语显示内容
- `kana`：假名
- `kanji`：汉字
- `chinese`：中文
- `remark`：备注

修改时不要破坏 JSON 的双引号、逗号和方括号。添加后续课程时，继续增加 Lesson 51、Lesson 52 等记录，再重新生成即可。

修改 JSON 后不需要重新编译 Python 程序或桌面应用，但需要重新生成普通 Excel 和 Word。

注意：现有宏版 `.xlsm` 内部保存了一份题库，修改 JSON 后不会自动同步。应先重新生成普通默写 `.xlsx`，再更新宏版后发布。

## 重新制作宏版（Windows）

先生成普通 Excel，然后运行：

```powershell
python main.py excel
./scripts/add_excel_macro.ps1 `
  -SourcePath "./output/大家的日语Ⅰ默写.xlsx" `
  -OutputPath "./output/大家的日语Ⅰ默写-宏版.xlsm"
```

该步骤需要本机安装 Microsoft Excel，并允许程序访问 VBA 工程。macOS 可以正常使用已经生成的宏版，但当前宏版制作脚本需要在 Windows 上运行。

## 构建桌面应用

不编译也可以使用源码。只有希望普通用户无需安装 Python、直接双击运行时，才需要构建桌面应用。

Windows：

```powershell
./build_windows_app.ps1
```

macOS：

```bash
chmod +x build_macos_app.sh
./build_macos_app.sh
```

构建完成后：

- Windows：`dist/windows/大家的日语词库-Windows/`
- macOS：`dist/macos/大家的日语词库-macOS/`

两个目录都会自动包含 `data/vocabulary.json`、空的 `output/` 和 README。应用不会把词库永久封装进程序，修改 JSON 不需要重新编译；也可以直接在界面中选择其他 JSON。

GitHub Actions 已拆分为两个独立工作流：

- `Build Windows app`
- `Build macOS app`

可以在 Actions 页面单独手动运行需要的平台；推送 `v*` 标签时两个工作流都会执行。Windows 和 macOS Artifact 各自只包含对应平台的完整发布包，不会混在同一目录中。

## 项目结构

```text
minna-no-nihongo-vocabulary/
├── .github/workflows/
│   ├── build-windows.yml
│   └── build-macos.yml
├── data/
│   └── vocabulary.json
├── output/
│   └── 大家的日语Ⅰ默写-宏版.xlsm
├── scripts/
│   ├── add_excel_macro.ps1
│   ├── excel_builder.py
│   ├── export_word.py
│   └── random_exam.py
├── desktop_app.py
├── main.py
├── requirements.txt
├── build_windows_app.ps1
└── build_macos_app.sh
```

## Git 更新

验证修改后执行：

```bash
git add .
git commit -m "更新词库工程"
git push
```

普通生成文件、编译文件和临时文件默认不提交。宏版 `.xlsm` 保留在仓库中，方便下载后直接使用。公开发布前请自行确认词库数据的版权许可。
