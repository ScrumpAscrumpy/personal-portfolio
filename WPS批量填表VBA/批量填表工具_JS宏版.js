// =============================================================================
// WPS 批量 Word 填表工具 v3.0 — JS宏版（WPS Mac 通用）
//
// 使用步骤：
//   1) 工具 → 开发工具 → 把本文件代码全选粘入 → Cmd+S 保存
//   2) 工具 → 运行宏 → 先跑【创建数据源模板】→ 填好右列 → 另存为 数据源.docx
//   3) 再跑【批量填写Word文档】→ 按提示三步选文件
// =============================================================================


// ─────────────────────────────────────────────────────────────────────────────
// 主功能：批量填写
// ─────────────────────────────────────────────────────────────────────────────
function 批量填写Word文档() {

    // 第一步：选数据源文档
    var fd1 = Application.FileDialog(3);
    fd1.Title = "第一步：选择数据源文档（填好内容的两列表格）";
    fd1.AllowMultiSelect = false;
    if (fd1.Show() != -1) return;
    var srcPath = fd1.SelectedItems.Item(1);

    // 读取数据源
    var map = readDataSource(srcPath);
    if (map == null) return;
    var keys = [];
    for (var k in map) { if (map[k] !== "") keys.push(k); }
    if (keys.length == 0) {
        alert("数据源里没找到有内容的字段。\n请先在数据源右列填写内容，再运行此宏。");
        return;
    }
    alert("读取数据源成功，共 " + keys.length + " 个字段：\n" + keys.join("、"));

    // 第二步：选目标 Word 文档（可多选）
    var fd2 = Application.FileDialog(3);
    fd2.Title = "第二步：选择要填写的 Word 文档（按住 Cmd 可多选）";
    fd2.AllowMultiSelect = true;
    if (fd2.Show() != -1) return;
    var nDoc = fd2.SelectedItems.Count;
    if (nDoc == 0) return;
    var targets = [];
    for (var i = 1; i <= nDoc; i++) targets.push(fd2.SelectedItems.Item(i));

    // 第三步：选输出目录
    var fd3 = Application.FileDialog(4);
    fd3.Title = "第三步：选择输出目录（填好的新文档存在这里，原文件不改）";
    if (fd3.Show() != -1) return;
    var outDir = fd3.SelectedItems.Item(1);

    // 逐个处理
    var ts = timestamp();
    var okN = 0, badN = 0, report = "";
    for (var j = 0; j < targets.length; j++) {
        var tp = targets[j];
        var stem = fileStem(tp);
        var outPath = joinPath(outDir, stem + "_已填写_" + ts + ".docx");
        var result = fillOneDoc(tp, map, keys, outPath);
        if (result.ok) {
            okN++;
            report += "  [OK] " + stem + "\n       → " + result.path + "\n";
        } else {
            badN++;
            report += "  [失败] " + stem + "\n       原因：" + result.err + "\n";
        }
    }

    alert("完成！成功 " + okN + " 份，失败 " + badN + " 份。\n\n" + report);
}


// ─────────────────────────────────────────────────────────────────────────────
// 创建数据源模板
// ─────────────────────────────────────────────────────────────────────────────
function 创建数据源模板() {
    var FIELDS = [
        "姓名", "性别", "户籍省市", "户籍地址", "手机号", "身份证号码",
        "最后缴费年月", "开户银行", "社保卡号",
        "家属名字", "家属身份证号", "家属电话"
    ];
    var doc = Application.Documents.Add();
    var r0 = doc.Range(0, 0);
    r0.InsertAfter(
        "数据源模板\n" +
        "在右列填好内容后，另存为【数据源.docx】。左列字段名勿改动。\n\n");

    var endPos = doc.Content.End - 1;
    var endRng = doc.Range(endPos, endPos);
    var tbl = doc.Tables.Add(endRng, FIELDS.length + 1, 2);
    tbl.Borders.Enable = true;

    tbl.Cell(1, 1).Range.Text = "字段名（对应文档中 {{字段名}} 的名称）";
    tbl.Cell(1, 2).Range.Text = "填写内容";
    for (var i = 0; i < FIELDS.length; i++) {
        tbl.Cell(i + 2, 1).Range.Text = FIELDS[i];
        tbl.Cell(i + 2, 2).Range.Text = "";
    }

    alert("数据源模板已创建！\n请在右列逐行填写，填完后点【文件→另存为】保存为数据源.docx。");
}


// ─────────────────────────────────────────────────────────────────────────────
// 读取数据源 → {字段名: 值}
// ─────────────────────────────────────────────────────────────────────────────
function readDataSource(path) {
    var doc = null, alreadyOpen = false;
    for (var d = 1; d <= Application.Documents.Count; d++) {
        if (Application.Documents.Item(d).FullName == path) {
            doc = Application.Documents.Item(d);
            alreadyOpen = true;
            break;
        }
    }
    if (!doc) doc = Application.Documents.Open(path);

    if (doc.Tables.Count == 0) {
        alert("数据源文档里没有表格！");
        if (!alreadyOpen) doc.Close(0);
        return null;
    }

    var map = {};
    var tbl = doc.Tables.Item(1);

    // 方案 A：按行 → 逐格读
    try {
        var rCount = tbl.Rows.Count;
        for (var r = 1; r <= rCount; r++) {
            var key = "", val = "";
            try {
                var cells = tbl.Rows.Item(r).Cells;
                if (cells.Count >= 1) key = cc(cells.Item(1).Range.Text);
                if (cells.Count >= 2) val = cc(cells.Item(2).Range.Text);
            } catch(e) { continue; }
            if (key && key.indexOf("字段名") !== 0) map[key] = val;
        }
    } catch(eA) {}

    // 方案 B 兜底：按 \x07 拆整表文本
    if (nKeys(map) == 0) {
        try {
            var parts = ("" + tbl.Range.Text).split("\x07");
            for (var p = 0; p + 1 < parts.length; p += 2) {
                var k2 = cc(parts[p]), v2 = cc(parts[p + 1]);
                if (k2 && k2.indexOf("字段名") !== 0) map[k2] = v2;
            }
        } catch(eB) {}
    }

    if (!alreadyOpen) doc.Close(0);
    return map;
}


// ─────────────────────────────────────────────────────────────────────────────
// 填写单个文档（返回 {ok, path, err}）
// ─────────────────────────────────────────────────────────────────────────────
function fillOneDoc(srcPath, map, keys, outPath) {
    var doc = null;
    try {
        doc = Application.Documents.Open(srcPath);

        // ── 占位符替换（两路方案）────────────────────────────────────────────
        for (var i = 0; i < keys.length; i++) {
            var key = keys[i];
            var val = "" + (map[key] != null ? map[key] : "");
            var ph  = "{{" + key + "}}";
            doReplace(doc, ph, val);
        }

        // ── 另存为新文件（依次尝试直到成功）─────────────────────────────────
        var saved = false, savedPath = "";
        var tries = [
            function() { doc.SaveAs2(outPath, 12); },  // 12 = wdFormatXMLDocument
            function() { doc.SaveAs2(outPath); },
            function() { doc.SaveAs(outPath, 12); },
            function() { doc.SaveAs(outPath); }
        ];
        for (var t = 0; t < tries.length; t++) {
            try { tries[t](); savedPath = doc.FullName; saved = true; break; } catch(e) {}
        }

        doc.Close(0);  // 关闭（另存已写盘，Close 不保存到原路径）

        if (!saved) return { ok: false, err: "SaveAs 全部失败，请检查输出目录写权限" };
        return { ok: true, path: savedPath };

    } catch(err) {
        try { if (doc) doc.Close(0); } catch(e) {}
        return { ok: false, err: err.message || String(err) };
    }
}


// ─────────────────────────────────────────────────────────────────────────────
// 文档内占位符替换（方案A: Find属性法  方案B: 逐格替换法）
// ─────────────────────────────────────────────────────────────────────────────
function doReplace(doc, ph, val) {
    // 方案 A：设置 Find 对象属性后调用 Execute
    var done = false;
    try {
        var find = doc.Content.Find;
        find.ClearFormatting();
        find.Replacement.ClearFormatting();
        find.Text = ph;
        find.Replacement.Text = val;
        find.Forward = true;
        find.Wrap = 1;          // 1 = wdFindContinue
        find.Format = false;
        find.MatchCase = false;
        find.MatchWholeWord = false;
        find.MatchWildcards = false;
        // 只传 Replace=2(wdReplaceAll)，其余用已设的属性
        find.Execute(ph, false, false, false, false, false, true, 1, false, val, 2);
        done = true;
    } catch(eA) {}

    // 方案 B：逐表格单元格文字替换
    if (!done) {
        var rx = new RegExp(
            ph.replace(/\{/g, "\\{").replace(/\}/g, "\\}"), "g");
        try {
            for (var t = 1; t <= doc.Tables.Count; t++) {
                var tbl = doc.Tables.Item(t);
                for (var r = 1; r <= tbl.Rows.Count; r++) {
                    var row = tbl.Rows.Item(r);
                    for (var c = 1; c <= row.Cells.Count; c++) {
                        try {
                            var cell = row.Cells.Item(c);
                            var txt = cc(cell.Range.Text);
                            if (txt.indexOf(ph) >= 0) {
                                var newTxt = txt.replace(rx, val);
                                var rng = cell.Range;
                                rng.End = rng.End - 1;
                                rng.Text = newTxt;
                            }
                        } catch(e2) {}
                    }
                }
            }
        } catch(eB) {}
    }
}


// ─────────────────────────────────────────────────────────────────────────────
// 工具函数
// ─────────────────────────────────────────────────────────────────────────────

// 清理单元格文本（去控制字符 + trim）
function cc(s) {
    if (s == null) return "";
    s = "" + s;
    s = s.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]/g, "");
    return s.replace(/^\s+|\s+$/g, "");
}

// 统计对象 key 数量
function nKeys(obj) {
    var n = 0;
    for (var k in obj) n++;
    return n;
}

// 拼路径（自动识别 Mac/Windows 分隔符）
function joinPath(dir, name) {
    var d = dir.replace(/[/\\]+$/, "");
    var sep = (d.indexOf("/") >= 0 || d.charAt(0) === "/") ? "/" : "\\";
    return d + sep + name;
}

// 提取文件名（不含扩展名）
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
