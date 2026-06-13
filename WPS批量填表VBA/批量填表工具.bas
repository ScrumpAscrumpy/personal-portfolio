Attribute VB_Name = "批量填表工具"
Option Explicit

'=============================================================================
' WPS 批量 Word 文档填写工具
' 版本：1.0  兼容 WPS Office（Windows / macOS）
'
' 功能：
'   1. 读取"数据源"Word文档（两列表格：左=字段名，右=填写内容）
'   2. 批量将字段内容写入多个目标 Word 文档对应位置
'   3. 输出带时间戳的副本，不破坏原始文件
'
' 使用方法：
'   在 WPS 宏编辑器中运行 【批量填写Word文档】 即可
'=============================================================================


' ─────────────────────────────────────────────────────────────────────────────
'  入口：批量填写
' ─────────────────────────────────────────────────────────────────────────────
Sub 批量填写Word文档()

    '── 第一步：选择数据源文档 ──────────────────────────────────────────────
    MsgBox "【第一步】即将选择数据源文档。" & vbCrLf & vbCrLf & _
           "数据源是一个 Word 文档，里面有一个两列表格：" & vbCrLf & _
           "  左列 = 字段名（如：姓名、身份证号）" & vbCrLf & _
           "  右列 = 要填写的内容（如：张三、110101...）" & vbCrLf & vbCrLf & _
           "如果没有数据源，请先运行【创建数据源模板】生成一个。", _
           vbInformation, "WPS批量填表工具"

    Dim dataPath As String
    dataPath = PickFile("选择数据源文档（含字段名和填写内容）")
    If dataPath = "" Then MsgBox "已取消操作。", vbInformation: Exit Sub

    '── 第二步：读取字段值 ──────────────────────────────────────────────────
    Dim fieldMap As Object
    Set fieldMap = ReadDataSource(dataPath)

    If fieldMap Is Nothing Or fieldMap.Count = 0 Then
        MsgBox "数据源文档中未找到任何字段。" & vbCrLf & vbCrLf & _
               "请确认数据源文档包含一个两列表格。", vbExclamation, "读取失败"
        Exit Sub
    End If

    '── 第三步：选择目标文档（多选）──────────────────────────────────────────
    MsgBox "【第二步】已读取 " & fieldMap.Count & " 个字段，现在选择要批量填写的目标 Word 文档。" & vbCrLf & _
           "（可按住 Ctrl 或 Shift 多选）", vbInformation, "WPS批量填表工具"

    Dim targets As Variant
    targets = PickMultipleFiles("选择目标 Word 文档（可多选）")
    If IsEmpty(targets) Then MsgBox "已取消操作。", vbInformation: Exit Sub

    '── 第四步：选择输出目录 ─────────────────────────────────────────────────
    Dim outDir As String
    outDir = PickFolder("【第三步】选择输出目录（存放填好的新文档）")
    If outDir = "" Then MsgBox "已取消操作。", vbInformation: Exit Sub

    '── 第五步：批量处理 ─────────────────────────────────────────────────────
    Dim ts As String
    ts = Format(Now(), "yyyymmdd_HHMMSS")

    Dim successCount As Integer: successCount = 0
    Dim failCount As Integer:    failCount = 0
    Dim failLog As String:       failLog = ""
    Dim i As Integer

    For i = 0 To UBound(targets)
        Dim srcPath As String: srcPath = targets(i)
        Dim baseName As String: baseName = FileBaseName(srcPath)
        Dim outName As String:  outName = baseName & "_已填写_" & ts & ".docx"
        Dim outPath As String:  outPath = outDir & AppSep() & outName

        Dim errMsg As String: errMsg = ""
        If FillDocument(srcPath, fieldMap, outPath, errMsg) Then
            successCount = successCount + 1
        Else
            failCount = failCount + 1
            failLog = failLog & vbCrLf & "  ✗ " & baseName & " — " & errMsg
        End If
    Next i

    '── 结果报告 ────────────────────────────────────────────────────────────
    Dim resultMsg As String
    resultMsg = "✅ 批量填写完成！" & vbCrLf & vbCrLf
    resultMsg = resultMsg & "成功：" & successCount & " 个文档" & vbCrLf
    If failCount > 0 Then
        resultMsg = resultMsg & "失败：" & failCount & " 个文档" & failLog & vbCrLf
    End If
    resultMsg = resultMsg & vbCrLf & "输出目录：" & vbCrLf & outDir

    MsgBox resultMsg, vbInformation, "WPS批量填表工具"
    OpenFolder outDir

End Sub


' ─────────────────────────────────────────────────────────────────────────────
'  创建数据源模板（首次使用时运行）
' ─────────────────────────────────────────────────────────────────────────────
Sub 创建数据源模板()

    Dim savePath As String
    savePath = PickFolder("选择数据源模板的保存位置")
    If savePath = "" Then Exit Sub
    savePath = savePath & AppSep() & "数据源_填写内容.docx"

    Dim doc As Document
    Set doc = Documents.Add

    '── 标题 ────────────────────────────────────────────────────────────────
    With doc.Paragraphs(1).Range
        .Text = "数据源 — 批量填写内容"
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Size = 16
        .Font.Bold = True
    End With

    doc.Range.InsertParagraphAfter

    '── 说明段 ───────────────────────────────────────────────────────────────
    Dim instrPara As Paragraph
    Set instrPara = doc.Paragraphs.Add
    With instrPara.Range
        .Text = "使用说明：在右列"填写内容"栏填入信息，左列字段名称与目标Word文档中的标签必须完全一致。"
        .Font.Size = 10
        .Font.Color = RGB(100, 100, 100)
    End With

    doc.Range.InsertParagraphAfter

    '── 数据表格（两列：字段名 | 填写内容）───────────────────────────────────
    Dim rng As Range
    Set rng = doc.Paragraphs.Last.Range

    Dim tbl As Table
    Dim fieldList As Variant
    fieldList = Array( _
        "姓名", "性别", "出生日期", "民族", "身份证号", "联系电话", _
        "户籍地址", "现居住地址", "工作单位", "职务", "学历", "政治面貌", _
        "婚姻状况", "退休日期", "社保编号", "参保单位", "银行账号", "备注" _
    )

    Set tbl = doc.Tables.Add(rng, UBound(fieldList) + 2, 2)
    tbl.Style = "Table Grid"

    '── 表头 ────────────────────────────────────────────────────────────────
    With tbl.Rows(1)
        .Cells(1).Range.Text = "字段名（与表格标签一致）"
        .Cells(2).Range.Text = "填写内容"
        .Cells(1).Range.Font.Bold = True
        .Cells(2).Range.Font.Bold = True
        .Cells(1).Shading.BackgroundPatternColor = RGB(189, 215, 238)
        .Cells(2).Shading.BackgroundPatternColor = RGB(189, 215, 238)
    End With

    '── 字段行 ───────────────────────────────────────────────────────────────
    Dim k As Integer
    For k = 0 To UBound(fieldList)
        tbl.Cell(k + 2, 1).Range.Text = fieldList(k)
        tbl.Cell(k + 2, 2).Range.Text = ""   ' 用户在此填写
    Next k

    '── 列宽 ────────────────────────────────────────────────────────────────
    tbl.Columns(1).Width = CentimetersToPoints(5)
    tbl.Columns(2).Width = CentimetersToPoints(10)

    '── 保存 ─────────────────────────────────────────────────────────────────
    doc.SaveAs2 Filename:=savePath, FileFormat:=wdFormatXMLDocument
    MsgBox "数据源模板已创建：" & vbCrLf & savePath & vbCrLf & vbCrLf & _
           "请在右列填入信息后，运行【批量填写Word文档】。", _
           vbInformation, "模板创建成功"

    OpenFolder (Left(savePath, InStrRev(savePath, AppSep()) - 1))

End Sub


' ─────────────────────────────────────────────────────────────────────────────
'  读取数据源文档 → 返回 Dictionary {字段名: 值}
' ─────────────────────────────────────────────────────────────────────────────
Function ReadDataSource(docPath As String) As Object

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1   ' 不区分大小写

    Dim doc As Document
    Dim wasOpen As Boolean
    wasOpen = IsDocOpen(docPath)

    On Error GoTo ErrOpen
    If wasOpen Then
        Set doc = GetOpenDoc(docPath)
    Else
        Set doc = Documents.Open(Filename:=docPath, ReadOnly:=True, Visible:=False)
    End If
    On Error GoTo 0

    '── 扫描所有表格，取两列行 ───────────────────────────────────────────────
    Dim tbl As Table
    For Each tbl In doc.Tables
        Dim rw As Row
        For Each rw In tbl.Rows
            Dim uc As Collection
            Set uc = UniqueCells(rw)
            If uc.Count >= 2 Then
                Dim lbl As String: lbl = CellText(uc(1))
                Dim val As String: val = CellText(uc(2))
                ' 字段名：非空、不超过20字、不含复选框符号
                If Len(lbl) >= 1 And Len(lbl) <= 20 Then
                    If InStr(lbl, "□") = 0 And InStr(lbl, "■") = 0 Then
                        If Not dict.Exists(lbl) Then
                            dict.Add lbl, val
                        End If
                    End If
                End If
            End If
        Next rw
    Next tbl

    If Not wasOpen Then doc.Close SaveChanges:=False
    Set ReadDataSource = dict
    Exit Function

ErrOpen:
    MsgBox "无法打开数据源文档：" & vbCrLf & docPath & vbCrLf & vbCrLf & _
           "错误：" & Err.Description, vbCritical
    Set ReadDataSource = Nothing

End Function


' ─────────────────────────────────────────────────────────────────────────────
'  填写单个目标文档
' ─────────────────────────────────────────────────────────────────────────────
Function FillDocument(srcPath As String, _
                      fieldMap As Object, _
                      outPath As String, _
                      ByRef errMsg As String) As Boolean

    On Error GoTo ErrHandler

    '── 复制原文件到输出路径（不修改原始文档）──────────────────────────────────
    FileCopy srcPath, outPath

    Dim doc As Document
    Set doc = Documents.Open(Filename:=outPath, Visible:=False)

    Dim applied As Object
    Set applied = CreateObject("Scripting.Dictionary")

    '── 扫描表格，匹配标签并写入 ────────────────────────────────────────────
    Dim tbl As Table
    For Each tbl In doc.Tables
        Dim rw As Row
        For Each rw In tbl.Rows
            Dim uc As Collection
            Set uc = UniqueCells(rw)
            Dim i As Integer
            For i = 1 To uc.Count - 1
                Dim labelText As String
                labelText = CellText(uc(i))

                If IsLabelCell(labelText) Then
                    Dim nextText As String
                    nextText = CellText(uc(i + 1))

                    If IsValueCell(nextText) Then
                        If fieldMap.Exists(labelText) And Not applied.Exists(labelText) Then
                            Dim v As String: v = fieldMap(labelText)
                            If v <> "" Then
                                WriteCell uc(i + 1), v
                                applied(labelText) = True
                            End If
                        End If
                    End If
                End If
            Next i
        Next rw
    Next tbl

    '── 扫描段落（"字段名：     "格式）──────────────────────────────────────
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        Dim paraText As String
        paraText = para.Range.Text
        If InStr(paraText, "：") > 0 Or InStr(paraText, ":") > 0 Then
            FillParagraphFields para, fieldMap, applied
        End If
    Next para

    doc.Save
    doc.Close
    FillDocument = True
    Exit Function

ErrHandler:
    errMsg = Err.Description
    FillDocument = False
    On Error Resume Next
    doc.Close SaveChanges:=False

End Function


' ─────────────────────────────────────────────────────────────────────────────
'  处理段落中 "标签：     " 格式的字段
' ─────────────────────────────────────────────────────────────────────────────
Sub FillParagraphFields(para As Paragraph, fieldMap As Object, applied As Object)
    Dim fullText As String
    fullText = para.Range.Text

    ' 匹配模式：标签名（1-15字）+ 冒号 + 3个以上空格
    Dim pattern As String
    pattern = "[\x{4e00}-\x{9fff}\w]{1,15}[：:]\s{3,}"

    ' 遍历 fieldMap 中的字段，逐一检查是否在段落中
    Dim key As Variant
    For Each key In fieldMap.Keys
        If Not applied.Exists(CStr(key)) Then
            Dim lbl As String: lbl = CStr(key)
            Dim val As String: val = CStr(fieldMap(lbl))
            If val = "" Then GoTo NextKey

            ' 查找 "字段名：   " 模式
            Dim colonPos As Long
            colonPos = InStr(fullText, lbl & "：")
            If colonPos = 0 Then colonPos = InStr(fullText, lbl & ":")
            If colonPos > 0 Then
                ' 找到字段名后的空格串，替换为值
                Dim afterColon As Long
                afterColon = colonPos + Len(lbl) + 1
                ' 检查后面是否有连续空格
                Dim spaceCount As Integer: spaceCount = 0
                Dim pos As Long: pos = afterColon
                Do While pos <= Len(fullText) And Mid(fullText, pos, 1) = " "
                    spaceCount = spaceCount + 1
                    pos = pos + 1
                Loop
                If spaceCount >= 3 Then
                    ' 替换段落文本
                    Dim newText As String
                    newText = Left(fullText, afterColon - 1) & val & Mid(fullText, afterColon + spaceCount)
                    para.Range.Text = newText
                    fullText = newText
                    applied(lbl) = True
                End If
            End If
        End If
NextKey:
    Next key
End Sub


' ─────────────────────────────────────────────────────────────────────────────
'  写入单元格（保留格式，只替换文字）
' ─────────────────────────────────────────────────────────────────────────────
Sub WriteCell(c As Cell, value As String)
    Dim rng As Range
    Set rng = c.Range

    ' 去掉末尾的段落符和单元格终止符，只替换文字内容
    rng.End = rng.End - 1   ' 不选中单元格终止符
    rng.Text = value
End Sub


' ─────────────────────────────────────────────────────────────────────────────
'  获取行中的唯一单元格（正确处理合并单元格）
'  与 Python 版本 _unique_cells() 逻辑完全对应
' ─────────────────────────────────────────────────────────────────────────────
Function UniqueCells(rw As Row) As Collection
    Dim result As New Collection
    Dim seenStarts() As Long
    ReDim seenStarts(rw.Cells.Count)
    Dim n As Integer: n = 0

    Dim c As Cell
    For Each c In rw.Cells
        Dim startPos As Long: startPos = c.Range.Start
        Dim found As Boolean: found = False
        Dim j As Integer
        For j = 0 To n - 1
            If seenStarts(j) = startPos Then found = True: Exit For
        Next j
        If Not found Then
            result.Add c
            seenStarts(n) = startPos
            n = n + 1
        End If
    Next c

    Set UniqueCells = result
End Function


' ─────────────────────────────────────────────────────────────────────────────
'  判断是否为标签单元格（对应 Python 版 _is_label()）
' ─────────────────────────────────────────────────────────────────────────────
Function IsLabelCell(text As String) As Boolean
    If Len(text) = 0 Or Len(text) > 16 Then
        IsLabelCell = False: Exit Function
    End If
    If InStr(text, "□") > 0 Or InStr(text, "■") > 0 Or InStr(text, "●") > 0 Then
        IsLabelCell = False: Exit Function
    End If
    IsLabelCell = True
End Function


' ─────────────────────────────────────────────────────────────────────────────
'  判断是否为可填写的空白单元格（对应 Python 版 _is_value_cell()）
' ─────────────────────────────────────────────────────────────────────────────
Function IsValueCell(text As String) As Boolean
    ' 完全为空
    If Len(text) = 0 Or Len(Trim(text)) = 0 Then
        IsValueCell = True: Exit Function
    End If
    ' 超过60字说明是正文，不是空白格
    If Len(text) > 60 Then
        IsValueCell = False: Exit Function
    End If
    ' 含复选框选项的不是空白格
    If InStr(text, "□") > 0 Or InStr(text, "■") > 0 Then
        IsValueCell = False: Exit Function
    End If
    ' 日期格式提示：如 "年   月   日"
    If InStr(text, "年") > 0 And InStr(text, "月") > 0 And Len(text) < 15 Then
        IsValueCell = True: Exit Function
    End If
    ' 地址格式提示：如 "省   市   县"
    If InStr(text, "省") > 0 And InStr(text, "市") > 0 And Len(text) < 25 Then
        IsValueCell = True: Exit Function
    End If
    IsValueCell = False
End Function


' ─────────────────────────────────────────────────────────────────────────────
'  清理单元格文字（去除 Word 内部的段落符和单元格结束符）
' ─────────────────────────────────────────────────────────────────────────────
Function CellText(c As Cell) As String
    Dim t As String
    t = c.Range.Text
    ' Word 单元格文字末尾固定带 Chr(13)+Chr(7)
    Do While Len(t) > 0 And (Right(t, 1) = Chr(13) Or Right(t, 1) = Chr(7))
        t = Left(t, Len(t) - 1)
    Loop
    CellText = Trim(t)
End Function


' ─────────────────────────────────────────────────────────────────────────────
'  文件 / 目录选择对话框（兼容 Windows / macOS）
' ─────────────────────────────────────────────────────────────────────────────
Function PickFile(title As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = title
        .Filters.Clear
        .Filters.Add "Word 文档", "*.docx;*.doc"
        .MultiSelect = False
        If .Show = -1 Then
            PickFile = .SelectedItems(1)
        Else
            PickFile = ""
        End If
    End With
End Function

Function PickMultipleFiles(title As String) As Variant
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = title
        .Filters.Clear
        .Filters.Add "Word 文档", "*.docx;*.doc"
        .MultiSelect = True
        If .Show = -1 Then
            Dim arr() As String
            ReDim arr(.SelectedItems.Count - 1)
            Dim k As Integer
            For k = 1 To .SelectedItems.Count
                arr(k - 1) = .SelectedItems(k)
            Next k
            PickMultipleFiles = arr
        Else
            PickMultipleFiles = Empty
        End If
    End With
End Function

Function PickFolder(title As String) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = title
        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
        Else
            PickFolder = ""
        End If
    End With
End Function

Sub OpenFolder(path As String)
    On Error Resume Next
    If AppSep() = "/" Then
        Shell "open """ & path & """"   ' macOS
    Else
        Shell "explorer.exe """ & path & """"   ' Windows
    End If
End Sub


' ─────────────────────────────────────────────────────────────────────────────
'  工具函数
' ─────────────────────────────────────────────────────────────────────────────
Function AppSep() As String
    AppSep = Application.PathSeparator
End Function

Function FileBaseName(path As String) As String
    Dim sep As String: sep = AppSep()
    Dim name As String
    name = Mid(path, InStrRev(path, sep) + 1)
    If InStr(name, ".") > 0 Then
        name = Left(name, InStrRev(name, ".") - 1)
    End If
    FileBaseName = name
End Function

Function IsDocOpen(path As String) As Boolean
    Dim d As Document
    For Each d In Documents
        If LCase(d.FullName) = LCase(path) Then
            IsDocOpen = True: Exit Function
        End If
    Next d
    IsDocOpen = False
End Function

Function GetOpenDoc(path As String) As Document
    Dim d As Document
    For Each d In Documents
        If LCase(d.FullName) = LCase(path) Then
            Set GetOpenDoc = d: Exit Function
        End If
    Next d
End Function
