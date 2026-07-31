param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

$vbaCode = @'
Option Explicit

Public Sub RegenerateExam()
    Dim wsSettings As Worksheet, wsSource As Worksheet, wsExam As Worksheet
    Dim lastRow As Long, rowIndex As Long, selectedCount As Long
    Dim requestedCount As Long, candidateCount As Long, outputCount As Long
    Dim candidates() As Long, tempValue As Long, swapIndex As Long
    Dim direction As String, lessonText As String, amountText As String
    Dim japanese As String, chinese As String

    Set wsSettings = ThisWorkbook.Worksheets("设置")
    Set wsSource = ThisWorkbook.Worksheets("题库")
    Set wsExam = ThisWorkbook.Worksheets("默写")

    direction = Trim$(CStr(wsSettings.Range("B2").Value))
    lessonText = NormalizeLessonText(CStr(wsSettings.Range("B3").Value))
    amountText = Trim$(CStr(wsSettings.Range("B4").Value))
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    ReDim candidates(1 To Application.Max(1, lastRow - 1))
    For rowIndex = 2 To lastRow
        If LessonIsSelected(CLng(wsSource.Cells(rowIndex, "A").Value), lessonText) Then
            candidateCount = candidateCount + 1
            candidates(candidateCount) = rowIndex
        End If
    Next rowIndex

    If candidateCount = 0 Then
        MsgBox "所选课程没有可用题目。", vbExclamation, "随机默写"
        Exit Sub
    End If

    ' 每次点击都用当前时间初始化随机数，再执行 Fisher-Yates 洗牌。
    Randomize
    For rowIndex = candidateCount To 2 Step -1
        swapIndex = Int(Rnd() * rowIndex) + 1
        tempValue = candidates(rowIndex)
        candidates(rowIndex) = candidates(swapIndex)
        candidates(swapIndex) = tempValue
    Next rowIndex

    If amountText = "全部" Or amountText = "" Then
        requestedCount = candidateCount
    Else
        requestedCount = CLng(Val(amountText))
    End If
    outputCount = Application.Min(requestedCount, candidateCount)

    ' 只更新已有题目区域，避免清空整张工作表造成等待。
    Application.ScreenUpdating = False
    Dim lastExamRow As Long
    lastExamRow = wsExam.Cells(wsExam.Rows.Count, "A").End(xlUp).Row
    If lastExamRow >= 2 Then wsExam.Range("A2:D" & lastExamRow).ClearContents

    For rowIndex = 1 To outputCount
        japanese = CStr(wsSource.Cells(candidates(rowIndex), "B").Value)
        chinese = CStr(wsSource.Cells(candidates(rowIndex), "C").Value)
        wsExam.Cells(rowIndex + 1, "A").Value = rowIndex
        If direction = "日译中" Then
            wsExam.Cells(rowIndex + 1, "B").Value = japanese
            wsExam.Cells(rowIndex + 1, "D").Value = chinese
        Else
            wsExam.Cells(rowIndex + 1, "B").Value = chinese
            wsExam.Cells(rowIndex + 1, "D").Value = japanese
        End If
    Next rowIndex

    Application.ScreenUpdating = True
    wsExam.Activate
    wsExam.Range("A2").Select
    MsgBox "已随机生成 " & outputCount & " 道题。", vbInformation, "随机默写"
End Sub

Private Function NormalizeLessonText(ByVal value As String) As String
    value = Replace(value, "，", ",")
    value = Replace(value, "、", ",")
    value = Replace(value, " ", "")
    NormalizeLessonText = value
End Function

Private Function LessonIsSelected(ByVal lesson As Long, ByVal lessonText As String) As Boolean
    If lessonText = "" Or lessonText = "全部" Then
        LessonIsSelected = True
    Else
        LessonIsSelected = InStr(1, "," & lessonText & ",", "," & CStr(lesson) & ",", vbTextCompare) > 0
    End If
End Function
'@

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open((Resolve-Path -LiteralPath $SourcePath).Path)

    $module = $workbook.VBProject.VBComponents.Add(1)
    $module.Name = "RandomExam"
    $module.CodeModule.AddFromString($vbaCode)

    $settings = $workbook.Worksheets.Item("设置")
    $button = $settings.Shapes.AddShape(5, 410, 34, 150, 42)
    $button.Name = "btnRegenerateExam"
    $button.TextFrame.Characters().Text = "随机生成"
    $button.OnAction = "RegenerateExam"
    $button.Fill.ForeColor.RGB = 8332469
    $button.Line.ForeColor.RGB = 8332469
    $button.TextFrame.Characters().Font.Color = 16777215
    $button.TextFrame.Characters().Font.Bold = $true

    $outputDirectory = Split-Path -Parent $OutputPath
    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory | Out-Null
    }
    $workbook.SaveAs($OutputPath, 52)
}
catch {
    Write-Error ("Macro workbook generation failed: " + $_.Exception.Message + [Environment]::NewLine + $_.ScriptStackTrace)
}
finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit() }
    if ($settings) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($settings) }
    if ($workbook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
    if ($excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
