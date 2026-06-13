// =============================================================================
// WPS 批量 Word 填表工具 — JS宏版（WPS Mac / Windows 通用）
//
// 用法：
//   1) 工具 → 开发工具（打开 JS宏编辑器）
//   2) 把本文件全部代码粘贴进去 → 保存
//   3) 工具 → 运行宏 → 选「批量填写Word文档」或「创建数据源模板」→ 运行
//
//   首次使用先跑【创建数据源模板】，右列填好信息另存为 数据源.docx，
//   再跑【批量填写Word文档】，三步选文件即可批量生成。
// =============================================================================

var FIELD_LIST = [
    "姓名", "性别", "户籍省市", "户籍地址", "手机号", "身份证号码",
    "最后缴费年月", "开户银行", "社保卡号",
    "家属名字", "家属身份证号", "家属电话"
];


// ─────────────────────────────────────────────────────────────────────────────
// 主功能：批量填写
// ─────────────────────────────────────────────────────────────────────────────
function 批量填写Word文档() {
    // 第一步：选数据源
    var fd1 = Application.FileDialog(3);          // 3 = msoFileDialogFilePicker
    fd1.Title = "第一步：选择数据源文档（两列表格：左=字段名，右=内容）";
    fd1.AllowMultiSelect = false;
    if (fd1.Show() != -1) return;
    var srcPath = fd1.SelectedItems.Item(1);

    // 读取数据源 → 字段字典
    var map = readDataSource(srcPath);
    if (map == null) return;
    var keys = [];
    for (var k in map) keys.push(k);
    if (keys.length == 0) {
        alert("数据源里没找到字段。请确认它含一个两列表格（左=字段名，右=内容）。");
        return;
    }

    // 第二步：选目标文档（可多选）
    var fd2 = Application.FileDialog(3);
    fd2.Title = "第二步：选择要填写的 Word 文档（按住 Cmd 多选）";
    fd2.AllowMultiSelect = true;
    if (fd2.Show() != -1) return;
    var nDoc = fd2.SelectedItems.Count;
    if (nDoc == 0) return;
    var targets = [];
    for (var i = 1; i <= nDoc; i++) targets.push(fd2.SelectedItems.Item(i));

    // 第三步：选输出目录
    var fd3 = Application.FileDialog(4);          // 4 = msoFileDialogFolderPicker
    fd3.Title = "第三步：选择输出目录（填好的新文档存这里，原文件不改）";
    if (fd3.Show() != -1) return;
    var outDir = fd3.SelectedItems.Item(1);

    // 逐个处理
    var ts = timestamp();
    var okN = 0, badN = 0, report = "";
    for (var j = 0; j < targets.length; j++) {
        var tp = targets[j];
        var stem = fileStem(tp);
        var outPath = joinPath(outDir, stem + "_已填写_" + ts + ".docx");
        try {
            fillOneDoc(tp, map, keys, outPath);
            okN++;
            report += "  [成功] " + stem + "\n";
        } catch (e) {
            badN++;
            report += "  [失败] " + stem + "  (" + e.message + ")\n";
        }
    }

    alert("完成！成功 " + okN + " 份，失败 " + badN + " 份。\n" +
          "保存位置：" + outDir + "\n\n" + report);
}


// ─────────────────────────────────────────────────────────────────────────────
// 创建数据源模板
// ─────────────────────────────────────────────────────────────────────────────
function 创建数据源模板() {
    var doc = Application.Documents.Add();
    var r0 = doc.Range(0, 0);
    r0.InsertAfter(
        "数据源模板：在右列填写内容，完成后【另存为 数据源.docx】。\n" +
        "左列字段名对应目标文档里 {{字段名}} 中的名称，请勿改动。\n\n");

    var n = FIELD_LIST.length;
    var endRng = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
    var tbl = doc.Tables.Add(endRng, n + 1, 2);
    tbl.Borders.Enable = true;

    tbl.Cell(1, 1).Range.Text = "字段名（勿改）";
    tbl.Cell(1, 2).Range.Text = "填写内容（在此输入）";
    for (var i = 0; i < n; i++) {
        tbl.Cell(i + 2, 1).Range.Text = FIELD_LIST[i];
        tbl.Cell(i + 2, 2).Range.Text = "";
    }

    alert("数据源模板已生成！\n请在右列逐行填写，完成后另存为【数据源.docx】。");
}


// ─────────────────────────────────────────────────────────────────────────────
// 读取数据源（两列表格）→ 返回 {字段名: 值}
// ─────────────────────────────────────────────────────────────────────────────
function readDataSource(path) {
    var doc = null, opened = false;
    // 若已打开则复用
    for (var d = 1; d <= Application.Documents.Count; d++) {
        if (Application.Documents.Item(d).FullName == path) {
            doc = Application.Documents.Item(d);
            opened = true;
            break;
        }
    }
    if (doc == null) {
        doc = Application.Documents.Open(path);
        opened = false;
    }

    if (doc.Tables.Count == 0) {
        alert("数据源文档里没有表格！");
        if (!opened) doc.Close(0);
        return null;
    }

    var map = {};
    var tbl = doc.Tables.Item(1);
    var rows = tbl.Rows.Count;
    for (var r = 1; r <= rows; r++) {
        var key = cleanCell(tbl.Cell(r, 1).Range.Text);
        var val = cleanCell(tbl.Cell(r, 2).Range.Text);
        if (key != "" && key.indexOf("字段名") != 0) {
            map[key] = val;
        }
    }

    if (!opened) doc.Close(0);   // 0 = 不保存
    return map;
}


// ─────────────────────────────────────────────────────────────────────────────
// 填写单个文档：打开原件 → 占位符替换 → 另存为副本（原件不改）
// ─────────────────────────────────────────────────────────────────────────────
function fillOneDoc(srcPath, map, keys, outPath) {
    var doc = Application.Documents.Open(srcPath);
    try {
        // 占位符替换 {{字段名}} → 值（覆盖正文+表格）
        for (var i = 0; i < keys.length; i++) {
            var key = keys[i];
            var val = map[key];
            var find = doc.Content.Find;
            find.ClearFormatting();
            find.Replacement.ClearFormatting();
            // Execute(FindText, MatchCase, WholeWord, Wildcards, SoundsLike,
            //         AllWordForms, Forward, Wrap(1), Format, ReplaceWith, Replace(2))
            find.Execute("{{" + key + "}}", false, false, false, false,
                         false, true, 1, false, val, 2);
        }
        // 另存为新文件（doc 指向新路径，关闭不影响原件）
        try { doc.SaveAs2(outPath, 12); }    // 12 = wdFormatXMLDocument(.docx)
        catch (e) { doc.SaveAs(outPath); }
        doc.Close(0);
    } catch (err) {
        try { doc.Close(0); } catch (e2) {}
        throw err;
    }
}


// ─────────────────────────────────────────────────────────────────────────────
// 工具函数
// ─────────────────────────────────────────────────────────────────────────────

// 去掉 Word 单元格文本末尾的控制字符(\r \x07)并 trim
function cleanCell(s) {
    if (s == null) return "";
    s = String(s);
    while (s.length > 0) {
        var c = s.charCodeAt(s.length - 1);
        if (c == 13 || c == 7 || c == 10) s = s.substring(0, s.length - 1);
        else break;
    }
    return s.replace(/^\s+|\s+$/g, "");
}

// 路径分隔符（Mac=/  Windows=\）
function pathSep() {
    var os = "" + Application.OperatingSystem;
    return (os.indexOf("Mac") >= 0 || os.indexOf("mac") >= 0) ? "/" : "\\";
}

function joinPath(dir, name) {
    var sep = pathSep();
    if (dir.charAt(dir.length - 1) == sep) return dir + name;
    return dir + sep + name;
}

// 取文件名（不含扩展名）
function fileStem(path) {
    var p = Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
    var name = (p >= 0) ? path.substring(p + 1) : path;
    var dot = name.lastIndexOf(".");
    return (dot > 0) ? name.substring(0, dot) : name;
}

// 时间戳 YYYYMMDD_HHmmss
function timestamp() {
    var d = new Date();
    function pad(n) { return (n < 10 ? "0" : "") + n; }
    return "" + d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) +
           "_" + pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds());
}
