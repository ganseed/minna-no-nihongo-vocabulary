# 大家的日语长期词库工程

本项目以 `data/vocabulary.json` 为唯一数据源，可以生成：

- 按课阅读的词库 Excel
- 支持选课、随机数量和日中方向的默写 Excel
- 词库 PDF
- 按实际题数排版的默写 PDF
- 可编辑 Word 词库

修改 JSON 后重新运行生成命令即可，不需要重新编译程序。

## 一、从 GitHub 下载

可以使用 Git 克隆：

```bash
git clone <你的 GitHub 仓库地址>
cd minna-no-nihongo-vocabulary
```

也可以在 GitHub 页面选择 `Code` → `Download ZIP`，下载后解压并进入项目目录。

## 二、安装运行环境

需要 Python 3.12 或更高版本。可使用下面的命令检查：

```bash
python --version
```

### Windows

安装 Python 时勾选 `Add Python to PATH`，然后在项目目录打开 PowerShell：

```powershell
python -m pip install -r requirements.txt
```

### macOS

在终端进入项目目录：

```bash
python3 --version
python3 -m pip install -r requirements.txt
```

安装依赖只需要执行一次。以后修改 JSON 后直接运行生成命令。

## 三、生成 Excel

Windows：

```powershell
python main.py excel
```

macOS：

```bash
python3 main.py excel
```

生成结果位于 `output/`：

- `大家的日语Ⅰ词库.xlsx`
- `大家的日语Ⅰ默写.xlsx`

默写 Excel 的“设置”工作表支持：

- B2：`日译中` 或 `中译日`
- B3：单课填 `3`；多课填 `1,3,5`；全部课程填 `全部`
- B4：`20`、`50`、`100`、`200` 或 `全部`

修改设置后 Excel 会重新计算题目。需要另一套随机顺序时，再运行一次 `excel` 命令。

## 四、导出词库 PDF 和 Word

Windows：

```powershell
python main.py pdf
python main.py word
```

macOS 请将 `python` 换成 `python3`。

生成文件位于 `output/`。

## 五、生成默写 PDF

默写 PDF 直接从 JSON 按实际题数生成，不会打印 Excel 中的预留空行，因此不会产生大量空白页。

从 Lesson 1、3、5 随机抽取 50 道日译中，并在末尾附答案：

```bash
python main.py exam-pdf --lessons 1,3,5 --count 50 --answers
```

从全部课程抽取 100 道中译日：

```bash
python main.py exam-pdf --count 100 --direction zh_to_jp
```

导出全部可默写词条：

```bash
python main.py exam-pdf --count all
```

固定随机题目，方便以后生成完全相同的一套卷子：

```bash
python main.py exam-pdf --lessons 1,2,3 --count 20 --seed 20260729 --answers
```

macOS 请将命令开头的 `python` 换成 `python3`。

## 六、修改词库

只需编辑：

```text
data/vocabulary.json
```

建议使用 VS Code，并保持 UTF-8 编码。每条记录的主要字段为：

- `id`：唯一编号，例如 `L01-001`
- `lesson`：课次
- `order`：本课顺序
- `type`：记录类型
- `japanese`：日语显示内容
- `kana`：假名
- `kanji`：汉字
- `chinese`：中文
- `remark`：备注

`type` 支持：

- `word`：单词，进入默写题库
- `expression`：固定表达，进入默写题库
- `example`：例句，不进入默认默写题库

修改时不要破坏 JSON 的引号、逗号和方括号。追加 Lesson26、Lesson27 等记录后，重新生成即可，不需要重做模板。

## 七、是否需要编译 EXE 或 Mac App

不需要。GitHub 源码安装 Python 和依赖后即可直接运行。

只有希望普通用户无需安装 Python、直接双击运行时，才需要另外发布：

- Windows `.exe`
- macOS `.app` 或原生可执行程序

这类编译文件建议放在 GitHub Releases，不要提交到源码仓库。

### 编译桌面应用

本项目提供同一套图形界面应用源码。应用不会把词库封装在程序内部，而是读取应用旁边的 `data/vocabulary.json`，因此修改 JSON 后不需要重新编译。

Windows 本机编译：

```powershell
./build_windows_app.ps1
```

输出位于 `dist/windows/`。

macOS 本机编译：

```bash
chmod +x build_macos_app.sh
./build_macos_app.sh
```

输出位于 `dist/macos/`。

也可以在 GitHub 仓库的 `Actions` 页面运行 `Build desktop apps`。工作流会分别使用 Windows 和 macOS 构建机生成两个安装包，在该次运行的 `Artifacts` 中下载：

- `windows-app`
- `macos-app`

发布包中应保持如下结构：

```text
应用程序或 App.app
data/vocabulary.json
output/
```

## 八、Git 更新流程

修改 JSON 或代码并验证生成正常后：

```bash
git add .
git commit -m "更新词库内容"
git push
```

教材 PDF、EXE、ZIP、OCR 临时文件和生成产物已通过 `.gitignore` 排除。公开发布前请自行确认词库数据与教材内容的版权许可。
