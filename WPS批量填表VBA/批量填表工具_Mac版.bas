Option Explicit
'=============================================================================
' WPS 批量 Word 文档填写工具 v2.1 — Mac/Windows 兼容版
'
' 运行【创建数据源模板】 → 填好右列 → 另存为 数据源.docx
' 运行【批量填写Word文档】→ 三步选文件 → 自动生成填好的副本
'=============================================================================

Const MAX_F As Integer = 200


'─────────────────────────────────────────────────────────────────────────────
' 主入口
'─────────────────────────────────────────────────────────────────────────────
Sub 批量填写Word文档()

    '── 第一步：选数据源 ──────────────────────────────────────────────────────
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "第一步：选择数据源文档（两列表格：左=字段名，右=填写内容）"
    fd.Filters.Clear
    fd.Filters.Add "Word 文档", "*.docx;*.doc"
    fd.AllowMultiSelect = False
    If fd.Show <> -1 Then Exit Sub
    Dim srcPath As String
    srcPath = fd.SelectedItems(1)

    '── 读取数据源 ─────────────────────────────────────────────────────────────
    Dim fKeys(MAX_F) As String, fVals(MAX_F) As String, fN As Integer
    fN = 0
    If Not ReadDS(srcPath, fKeys, fVals, fN) Then Exit Sub
    If fN = 0 Then
        MsgBox "数据源未找到任何字段，请确认包含两列表格（左=字段名，右=内容）。", vbExclamation
        Exit Sub
    End If

    '── 第二步：选目标 Word 文档（可多选）──────────────────────────────────────
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "第二步：选择要填写的 Word 文档（按住 Cmd 可多选）"
    fd.Filters.Clear
    fd.Filters.Add "Word 文档", "*.docx;*.doc"
    fd.AllowMultiSelect = True
    If fd.Show <> -1 Then Exit Sub
    If fd.SelectedItems.Count = 0 Then Exit Sub

    '── 第三步：选输出目录 ────────────────────────────────────────────────────
    Dim fdDir As FileDialog
    Set fdDir = Application.FileDialog(msoFileDialogFolderPicker)
    fdDir.Title = "第三步：选择输出目录（填好的新文档存这里，原文件不会被修改）"
    If fdDir.Show <> -1 Then Exit Sub
    Dim outDir As String
    outDir = fdDir.SelectedItems(1)

    '── 逐个处理目标文档 ──────────────────────────────────────────────────────
    Dim ts As String:   ts = Format(Now, "YYYYMMDD_HHmmss")
    Dim sep As String:  sep = Application.PathSeparator
    Dim okN As Integer: okN = 0
    Dim badN As Integer: badN = 0
    Dim rpt As String:  rpt = ""
    Dim i As Integer

    For i = 1 To fd.SelectedItems.Count
        Dim tp As String:   tp = fd.SelectedItems(i)
        Dim stem As String: stem = FileStem(tp)
        Dim op As String:   op = outDir & sep & stem & "_已填写_" & ts & ".docx"
        Dim em As String:   em = ""

        If FillOneDoc(tp, fKeys, fVals, fN, op, em) Then
            okN = okN + 1
            rpt = rpt & "  [OK] " & stem & vbCrLf
        Else
            badN = badN + 1
            rpt = rpt & "  [失败] " & stem & "  (" & em & ")" & vbCrLf
        End If
    Next i

    MsgBox "完成！成功 " & okN & " 份，失败 " & badN & " 份。" & vbCrLf & _
           "保存位置：" & outDir & vbCrLf & vbCrLf & rpt, _
           vbInformation, "批量填写完成"
    OpenFolder outDir

End Sub


'─────────────────────────────────────────────────────────────────────────────
' 生成数据源模板
'─────────────────────────────────────────────────────────────────────────────
Sub 创建数据源模板()
    Dim fl As Variant
    fl = Array("姓名", "性别", "户籍省市", "户籍地址", "手机号", "身份证号码", _
               "最后缴费年月", "开户银行", "社保卡号", _
               "家属名字", "家属身份证号", "家属电话")

    Dim doc As Document
    Set doc = Documents.Add

    ' 写说明段落
    doc.Paragraphs(1).Range.Text = "数据源模板"
    doc.Paragraphs(1).Range.InsertParagraphAfter
    doc.Paragraphs(2).Range.Text = _
        "填好右列内容后，另存为【数据源.docx】。" & _
        "字段名对应目标文档里 {{字段名}} 中的名称。"
    doc.Paragraphs(2).Range.InsertParagraphAfter

    ' 添加表格
    Dim tbl As Table
    Set tbl = doc.Tables.Add( _
        Range:=doc.Paragraphs(doc.Paragraphs.Count).Range, _
        NumRows:=UBound(fl) + 2, _
        NumColumns:=2)

    tbl.Cell(1, 1).Range.Text = "字段名（对应目标文档 {{}} 中的名称）"
    tbl.Cell(1, 2).Range.Text = "填写内容（在这里输入）"

    Dim j As Integer
    For j = 0 To UBound(fl)
        tbl.Cell(j + 2, 1).Range.Text = CStr(fl(j))
        tbl.Cell(j + 2, 2).Range.Text = ""
    Next j

    tbl.Borders.Enable = True

    MsgBox "数据源模板已创建！" & vbCrLf & _
           "请在右列逐行填写对应内容，完成后另存为【数据源.docx】。", _
           vbInformation, "提示"
End Sub


'─────────────────────────────────────────────────────────────────────────────
' 读取数据源文档 → 填充 fKeys/fVals 数组
'─────────────────────────────────────────────────────────────────────────────
Function ReadDS(path As String, fKeys() As String, fVals() As String, _
               ByRef fN As Integer) As Boolean
    On Error GoTo Err1
    fN = 0

    Dim doc As Document, wasOpen As Boolean
    wasOpen = False
    Dim d As Document
    For Each d In Documents
        If d.FullName = path Then
            Set doc = d
            wasOpen = True
            Exit For
        End If
    Next d
    If Not wasOpen Then
        Set doc = Documents.Open(Filename:=path, Visible:=False)
    End If

    If doc.Tables.Count = 0 Then
        MsgBox "数据源文档没有表格！", vbExclamation
        If Not wasOpen Then doc.Close False
        ReadDS = False
        Exit Function
    End If

    Dim tbl As Table
    Set tbl = doc.Tables(1)
    Dim r As Integer
    For r = 1 To tbl.Rows.Count
        Dim k As String: k = CellStr(tbl.Cell(r, 1))
        Dim v As String: v = CellStr(tbl.Cell(r, 2))
        ' 跳过表头行和空行
        If k <> "" And k <> "字段名" And _
           Left(k, 3) <> "字段名" And k <> "字段名（对应目标文档 {{}} 中的名称）" Then
            fKeys(fN) = k
            fVals(fN) = v
            fN = fN + 1
            If fN >= MAX_F Then Exit For
        End If
    Next r

    If Not wasOpen Then doc.Close False
    ReadDS = True
    Exit Function
Err1:
    MsgBox "读取数据源失败：" & Err.Description, vbCritical
    ReadDS = False
End Function


'─────────────────────────────────────────────────────────────────────────────
' 填写单个文档（先占位符替换，再智能识别空白格）
'─────────────────────────────────────────────────────────────────────────────
Function FillOneDoc(srcPath As String, fKeys() As String, fVals() As String, _
                   fN As Integer, outPath As String, ByRef em As String) As Boolean
    On Error GoTo ErrH

    Dim doc As Document
    Dim docOpened As Boolean
    docOpened = False

    FileCopy srcPath, outPath
    Set doc = Documents.Open(Filename:=outPath, Visible:=False)
    docOpened = True

    ' ── 主路径：占位符替换 {{字段名}} → 填写内容 ────────────────────────────
    DoReplace doc, fKeys, fVals, fN

    ' ── 备用路径：标签→空白格 智能识别（无占位符的普通表格） ────────────────
    SmartFill doc, fKeys, fVals, fN

    doc.Save
    doc.Close
    docOpened = False
    FillOneDoc = True
    Exit Function
ErrH:
    em = Err.Description
    FillOneDoc = False
    On Error Resume Next
    If docOpened Then doc.Close False
End Function


'─────────────────────────────────────────────────────────────────────────────
' 占位符替换（全文档，含表格）
'─────────────────────────────────────────────────────────────────────────────
Sub DoReplace(doc As Document, fKeys() As String, fVals() As String, fN As Integer)
    Dim i As Integer
    For i = 0 To fN - 1
        With doc.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text             = "{{" & fKeys(i) & "}}"
            .Replacement.Text = fVals(i)
            .Forward          = True
            .Wrap             = wdFindContinue
            .MatchWildcards   = False
            .MatchCase        = False
            .Execute Replace:=wdReplaceAll
        End With
    Next i
End Sub


'─────────────────────────────────────────────────────────────────────────────
' 标签→空白格 智能识别（备用，适用于无 {{}} 标记的普通表格）
'─────────────────────────────────────────────────────────────────────────────
Sub SmartFill(doc As Document, fKeys() As String, fVals() As String, fN As Integer)
    Dim tbl As Table
    For Each tbl In doc.Tables
        Dim rw As Row
        For Each rw In tbl.Rows
            Dim ci As Integer
            For ci = 1 To rw.Cells.Count - 1
                Dim lbl As String: lbl = CellStr(rw.Cells(ci))
                If IsLabel(lbl) Then
                    Dim nxt As String: nxt = CellStr(rw.Cells(ci + 1))
                    If IsBlankCell(nxt) Then
                        Dim idx As Integer: idx = FindKey(fKeys, fN, lbl)
                        If idx >= 0 And fVals(idx) <> "" Then
                            WriteCell rw.Cells(ci + 1), fVals(idx)
                        End If
                    End If
                End If
            Next ci
        Next rw
    Next tbl
End Sub


'─────────────────────────────────────────────────────────────────────────────
' 辅助函数
'─────────────────────────────────────────────────────────────────────────────

' 读取单元格文本，去掉 Word 末尾的 Chr(13)+Chr(7) 控制字符
Function CellStr(c As Cell) As String
    Dim s As String: s = c.Range.Text
    Do While Len(s) > 0
        Dim b As Integer: b = Asc(Right(s, 1))
        If b = 13 Or b = 7 Then
            s = Left(s, Len(s) - 1)
        Else
            Exit Do
        End If
    Loop
    CellStr = Trim(s)
End Function

' 判断是否标签格（1-16字，无复选框符号）
Function IsLabel(s As String) As Boolean
    If Len(s) = 0 Or Len(s) > 16 Then IsLabel = False: Exit Function
    If InStr(s, "□") > 0 Or InStr(s, "■") > 0 Then IsLabel = False: Exit Function
    IsLabel = True
End Function

' 判断是否空白值格（空，或仅含年月省市格式提示）
Function IsBlankCell(s As String) As Boolean
    If Len(s) = 0 Then IsBlankCell = True: Exit Function
    If Len(s) > 60 Then IsBlankCell = False: Exit Function
    If InStr(s, "□") > 0 Or InStr(s, "■") > 0 Then IsBlankCell = False: Exit Function
    Dim t As String: t = s
    Dim sub_ As Variant
    For Each sub_ In Array("年", "月", "日", "省", "市", "县", "区", " ", "　", Chr(32))
        t = Replace(t, CStr(sub_), "")
    Next sub_
    IsBlankCell = (Trim(t) = "")
End Function

' 写入单元格（不破坏表格结构）
Sub WriteCell(c As Cell, val As String)
    Dim rng As Range
    Set rng = c.Range
    rng.End = rng.End - 1   ' 排除单元格结束标记 Chr(13)
    rng.Text = val
End Sub

' 在 fKeys 数组中查找，返回下标（找不到返回 -1）
Function FindKey(fKeys() As String, fN As Integer, target As String) As Integer
    Dim i As Integer
    For i = 0 To fN - 1
        If fKeys(i) = target Then FindKey = i: Exit Function
    Next i
    FindKey = -1
End Function

' 提取路径中的文件名（不含扩展名）
Function FileStem(path As String) As String
    Dim sep As String: sep = Application.PathSeparator
    Dim p As Integer: p = InStrRev(path, sep)
    Dim name_ As String
    If p > 0 Then name_ = Mid(path, p + 1) Else name_ = path
    Dim dot As Integer: dot = InStrRev(name_, ".")
    If dot > 1 Then FileStem = Left(name_, dot - 1) Else FileStem = name_
End Function

' 打开文件夹（Mac 用 open，Windows 用 explorer）
Sub OpenFolder(path As String)
    On Error Resume Next
    If InStr(1, Application.OperatingSystem, "Mac") > 0 Or _
       InStr(1, Application.OperatingSystem, "mac") > 0 Then
        Shell "open " & Chr(34) & path & Chr(34)
    Else
        Shell "explorer.exe " & Chr(34) & path & Chr(34)
    End If
End Sub
