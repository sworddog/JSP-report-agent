<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.sql.*,java.util.*,java.net.*,java.text.*" %>
<%
    // ==================== 获取当前日期 ====================
    SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
    String todayStr = dateFormat.format(new java.util.Date());

    // ==================== 处理 POST 保存请求 ====================
    request.setCharacterEncoding("UTF-8");
    String action = request.getParameter("action");
    String saveMsg = "";
    boolean saveSuccess = false;

    if ("save".equals(action)) {
        String[] selectedIds = request.getParameterValues("selectedIds");
        if (selectedIds != null && selectedIds.length > 0) {
            Connection saveConn = null;
            PreparedStatement savePstmt = null;
            int savedCount = 0;
            try {
                Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
                String urlEcology = "jdbc:sqlserver://172.16.5.188:1433;databaseName=ecology";
                saveConn = DriverManager.getConnection(urlEcology, "sa", "Sble123456");

                // 先查询已存在的 fentryid，避免重复插入报错
                Set<String> existingIds = new HashSet<String>();
                PreparedStatement checkPstmt = saveConn.prepareStatement(
                    "SELECT fentryid FROM uf_wwwhskjl WHERE fentryid IN (" +
                    String.join(",", Collections.nCopies(selectedIds.length, "?")) + ")"
                );
                for (int i = 0; i < selectedIds.length; i++) {
                    checkPstmt.setInt(i + 1, Integer.parseInt(selectedIds[i]));
                }
                ResultSet checkRs = checkPstmt.executeQuery();
                while (checkRs.next()) {
                    existingIds.add(String.valueOf(checkRs.getInt("fentryid")));
                }
                checkRs.close();
                checkPstmt.close();

                // 插入或更新
                String upsertSql = "IF EXISTS (SELECT 1 FROM uf_wwwhskjl WHERE fentryid = ?) " +
                                   "UPDATE uf_wwwhskjl SET czrq = ?, sfsk = '1' WHERE fentryid = ? " +
                                   "ELSE " +
                                   "INSERT INTO uf_wwwhskjl (fentryid, czrq, sfsk) VALUES (?, ?, '1')";
                savePstmt = saveConn.prepareStatement(upsertSql);
                for (String idStr : selectedIds) {
                    int fid = Integer.parseInt(idStr);
                    savePstmt.setInt(1, fid);
                    savePstmt.setString(2, todayStr);
                    savePstmt.setInt(3, fid);
                    savePstmt.setInt(4, fid);
                    savePstmt.setString(5, todayStr);
                    savePstmt.executeUpdate();
                    savedCount++;
                }
                saveSuccess = true;
                saveMsg = "成功标注 " + savedCount + " 条记录为已收款";
            } catch (Exception e) {
                saveMsg = "保存失败：" + e.getMessage();
            } finally {
                try { if (savePstmt != null) savePstmt.close(); } catch (Exception e) {}
                try { if (saveConn != null) saveConn.close(); } catch (Exception e) {}
            }
        } else {
            saveMsg = "请至少选择一行数据";
        }
    }

    // ==================== 获取筛选参数 ====================
    String filterDateFrom = request.getParameter("dateFrom");
    if (filterDateFrom == null) filterDateFrom = "";
    filterDateFrom = filterDateFrom.trim();

    String filterDateTo = request.getParameter("dateTo");
    if (filterDateTo == null) filterDateTo = "";
    filterDateTo = filterDateTo.trim();

    String filterOrg = request.getParameter("orgFilter");
    if (filterOrg == null) filterOrg = "";
    filterOrg = filterOrg.trim();

    String filterReceipt = request.getParameter("receiptFilter");
    if (filterReceipt == null) filterReceipt = "";
    filterReceipt = filterReceipt.trim();

    boolean hasDateFilter = !filterDateFrom.isEmpty() || !filterDateTo.isEmpty();
    boolean hasOrgFilter = !filterOrg.isEmpty();
    boolean hasReceiptFilter = !filterReceipt.isEmpty();
    boolean hasFilter = hasDateFilter || hasOrgFilter || hasReceiptFilter;

    // ==================== 导出模式判断 ====================
    boolean isExport = "excel".equals(request.getParameter("export"));
    if (isExport) {
        response.setContentType("application/vnd.ms-excel; charset=UTF-8");
        response.setHeader("Content-Disposition", "attachment; filename=" +
            URLEncoder.encode("供应商未回收款核查报表.xls", "UTF-8"));
    }
%>
<% if (!isExport) { %>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>供应商未回收款核查报表</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px 30px; }
        h1 { color: #333; font-size: 20px; }
        .top-bar { display: flex; justify-content: space-between; align-items: center;
                   margin-bottom: 12px; flex-wrap: wrap; gap: 10px; }
        .top-bar h1 { margin: 0; }
        .btn-row { display: flex; gap: 10px; align-items: center; }
        .btn { padding: 7px 18px; border: none; border-radius: 4px; cursor: pointer;
               font-size: 14px; font-weight: bold; color: #fff; text-decoration: none; display: inline-block; }
        .btn-confirm { background: #e65100; }
        .btn-confirm:hover { background: #bf360c; }
        .btn-confirm:disabled { background: #ccc; cursor: not-allowed; }
        .btn-export { background: #28a745; }
        .btn-export:hover { background: #1e7e34; }
        .btn-clear { background: #999; }
        .btn-clear:hover { background: #777; }
        .msg { padding: 8px 16px; border-radius: 4px; margin-bottom: 12px; font-size: 14px; }
        .msg-success { background: #e8f5e9; color: #2e7d32; border: 1px solid #a5d6a7; }
        .msg-error { background: #ffebee; color: #c62828; border: 1px solid #ef9a9a; }
        .msg-info { background: #e3f2fd; color: #1565c0; border: 1px solid #90caf9; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; font-size: 13px; }
        th, td { border: 1px solid #ccc; padding: 6px 8px; text-align: left; }
        th { background-color: #f5f5f5; white-space: nowrap; position: sticky; top: 0; }
        .num { text-align: right; }
        .chk-col { text-align: center; width: 40px; }
        .error { color: red; }
        tr.row-marked { background-color: #f0f0f0; color: #aaa; }
        tr.row-marked td { text-decoration: line-through; text-decoration-color: #ccc; }
        .marked-tag { display: inline-block; background: #e8f5e9; color: #2e7d32;
                      padding: 2px 8px; border-radius: 10px; font-size: 11px; }
        .stats { display: flex; gap: 16px; margin-bottom: 12px; }
        .stat-item { background: #f9f9f9; padding: 8px 16px; border-radius: 4px; border: 1px solid #e0e0e0; }
        .stat-num { font-size: 20px; font-weight: bold; }
        .stat-label { font-size: 12px; color: #888; }
        .stat-item.warn .stat-num { color: #e65100; }
        .stat-item.ok .stat-num { color: #2e7d32; }
    </style>
</head>
<body>

    <div class="top-bar">
        <h1>供应商未回收款核查报表</h1>
        <div class="btn-row">
            <span style="color:#888;font-size:13px;margin-right:4px;">
                操作日期：<b><%= todayStr %></b>
            </span>
            <button class="btn btn-confirm" id="btnConfirm" onclick="doSave()" disabled>
                ✔ 确认收款
            </button>
            <a class="btn btn-export" href="<%= request.getRequestURI() %>?export=excel">导出Excel</a>
        </div>
    </div>

    <%-- 提示消息 --%>
    <% if (!saveMsg.isEmpty()) { %>
        <div class="msg <%= saveSuccess ? "msg-success" : "msg-error" %>">
            <%= saveMsg %>
        </div>
    <% } %>

    <%-- 筛选栏 --%>
    <div style="margin-bottom: 12px; padding: 10px 14px; background: #f9f9f9; border: 1px solid #e0e0e0; border-radius: 4px;">
        <form method="get" action="" style="display: flex; align-items: center; gap: 12px; flex-wrap: wrap;">
            <label>发出日期从：<input type="date" name="dateFrom" value="<%= filterDateFrom %>" style="padding: 4px 8px; border: 1px solid #ccc; border-radius: 3px; font-size: 13px;" /></label>
            <label>至：<input type="date" name="dateTo" value="<%= filterDateTo %>" style="padding: 4px 8px; border: 1px solid #ccc; border-radius: 3px; font-size: 13px;" /></label>
            <label>生产组织：
                <select name="orgFilter" style="padding: 4px 8px; border: 1px solid #ccc; border-radius: 3px; font-size: 13px;">
                    <option value="">全部</option>
                    <option value="0" <%= "0".equals(filterOrg) ? "selected" : "" %>>杭州</option>
                    <option value="1" <%= "1".equals(filterOrg) ? "selected" : "" %>>安吉</option>
                </select>
            </label>
            <label>是否收款：
                <select name="receiptFilter" style="padding: 4px 8px; border: 1px solid #ccc; border-radius: 3px; font-size: 13px;">
                    <option value="">全部</option>
                    <option value="1" <%= "1".equals(filterReceipt) ? "selected" : "" %>>是（已收款）</option>
                    <option value="0" <%= "0".equals(filterReceipt) ? "selected" : "" %>>否（待处理）</option>
                </select>
            </label>
            <button type="submit" style="padding: 5px 14px; background: #007acc; color: #fff; border: none; border-radius: 3px; cursor: pointer;">查询</button>
            <% if (hasFilter) { %>
                <a href="<%= request.getRequestURI() %>" style="padding: 5px 14px; background: #999; color: #fff; text-decoration: none; border-radius: 3px;">清除</a>
            <% } %>
        </form>
    </div>

    <%-- 统计栏 --%>
    <div class="stats" id="statsBar" style="display:none;">
        <div class="stat-item warn"><div class="stat-num" id="statPending">-</div><div class="stat-label">待处理（未回）</div></div>
        <div class="stat-item ok"><div class="stat-num" id="statDone">-</div><div class="stat-label">已确认收款</div></div>
        <div class="stat-item"><div class="stat-num" id="statTotal">-</div><div class="stat-label">总计</div></div>
    </div>

    <%-- 操作提示 --%>
    <div class="msg msg-info" style="font-size:13px;">
        💡 勾选行前方的复选框，点击右上角 <b>"确认收款"</b> 按钮，即可将选中行标注为已收款。
        <span style="color:#888;">灰色行表示已确认收款，不可重复操作。</span>
    </div>

<% } %>

<%
    // ==================== 查询金蝶视图（AISmonth6cs） ====================
    Connection connKingdee = null;
    Statement stmtKingdee = null;
    ResultSet rsKingdee = null;

    // 所有列名（按金蝶视图 JTgyswhmx 顺序）
    String[] colNames = {"FENTRYID","生产组织","发出日期","单据编号","物料编码","规格型号","供应商","生产订单状态","未回数量"};
    // 显示列名（FENTRYID 作为隐藏标识，不单独显示）
    String[] displayCols = {"生产组织","发出日期","单据编号","物料编码","规格型号","供应商","生产订单状态","未回数量"};

    List<Map<String, String>> kingdeeData = new ArrayList<Map<String, String>>();

    try {
        Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");

        String urlKd = "jdbc:sqlserver://172.16.5.185:1433;databaseName=AISmonth6cs";
        connKingdee = DriverManager.getConnection(urlKd, "onlyreaduser", "Supror@2003");
        stmtKingdee = connKingdee.createStatement();
        String sqlKd = "SELECT [FENTRYID],[生产组织],[发出日期],[单据编号],[物料编码],[规格型号],[供应商],[生产订单状态],[未回数量] " +
                       "FROM JTgyswhmx ORDER BY [发出日期] DESC";
        rsKingdee = stmtKingdee.executeQuery(sqlKd);

        while (rsKingdee.next()) {
            Map<String, String> row = new LinkedHashMap<String, String>();
            for (String col : colNames) {
                row.put(col, rsKingdee.getString(col));
            }
            kingdeeData.add(row);
        }
    } catch (Exception e) {
        if (!isExport) {
            out.println("<p class='error'>金蝶数据库（AISmonth6cs）连接或查询出错：" + e.getMessage() + "</p>");
        }
    } finally {
        try { if (rsKingdee != null) rsKingdee.close(); } catch (Exception e) {}
        try { if (stmtKingdee != null) stmtKingdee.close(); } catch (Exception e) {}
        try { if (connKingdee != null) connKingdee.close(); } catch (Exception e) {}
    }

    // ==================== 查询 OA 记录表（ecology） ====================
    Set<Integer> markedIds = new HashSet<Integer>();
    Map<Integer, String> markedDate = new HashMap<Integer, String>();

    Connection connOa = null;
    Statement stmtOa = null;
    ResultSet rsOa = null;

    try {
        Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");

        String urlOa = "jdbc:sqlserver://172.16.5.188:1433;databaseName=ecology";
        connOa = DriverManager.getConnection(urlOa, "sa", "Sble123456");
        stmtOa = connOa.createStatement();
        // 只查询 sfsk='1' 的记录
        String sqlOa = "SELECT fentryid, czrq FROM uf_wwwhskjl WHERE sfsk = '1'";
        rsOa = stmtOa.executeQuery(sqlOa);

        while (rsOa.next()) {
            int fid = rsOa.getInt("fentryid");
            markedIds.add(fid);
            String czrq = rsOa.getString("czrq");
            if (czrq != null) {
                markedDate.put(fid, czrq);
            }
        }
    } catch (Exception e) {
        if (!isExport) {
            out.println("<p class='error'>OA数据库（ecology）连接或查询出错：" + e.getMessage() + "</p>");
        }
    } finally {
        try { if (rsOa != null) rsOa.close(); } catch (Exception e) {}
        try { if (stmtOa != null) stmtOa.close(); } catch (Exception e) {}
        try { if (connOa != null) connOa.close(); } catch (Exception e) {}
    }
%>

    <form id="mainForm" method="post" action="">
        <input type="hidden" name="action" value="save" />
        <div id="selectedIdsContainer"></div>
    </form>

    <table id="dataTable">
        <thead>
            <tr>
                <th class="chk-col">
                    <% if (!isExport) { %>
                        <input type="checkbox" id="selectAll" title="全选/取消全选" onclick="toggleAll(this)" />
                    <% } else { %>
                        已收款
                    <% } %>
                </th>
                <% for (String col : displayCols) { %>
                    <th><%= col %></th>
                <% } %>
                <th>收款状态</th>
            </tr>
        </thead>
        <tbody>
<%
    int totalCount = kingdeeData.size();
    int markedCount = 0;
    int displayedCount = 0;
    int displayedMarked = 0;
    int rowIndex = 0;

    for (Map<String, String> row : kingdeeData) {
        rowIndex++;
        String fentryidStr = row.get("FENTRYID");
        int fid = 0;
        if (fentryidStr != null && !fentryidStr.trim().isEmpty()) {
            try { fid = Integer.parseInt(fentryidStr.trim()); } catch (NumberFormatException e) {}
        }
        boolean isMarked = markedIds.contains(fid);

        // ---- 筛选：发出日期区间 ----
        String issueDate = row.get("发出日期");
        if (issueDate == null) issueDate = "";
        issueDate = issueDate.trim();
        // 日期可能是 "2026-05-01" 或 "2026-05-01 00:00:00"，统一取前10位
        String issueDate10 = issueDate.length() >= 10 ? issueDate.substring(0, 10) : issueDate;
        if (hasDateFilter) {
            if (!filterDateFrom.isEmpty() && issueDate10.compareTo(filterDateFrom) < 0) {
                continue;
            }
            if (!filterDateTo.isEmpty() && issueDate10.compareTo(filterDateTo) > 0) {
                continue;
            }
        }

        // ---- 筛选：生产组织 ----
        String orgVal = row.get("生产组织");
        if (orgVal == null) orgVal = "";
        orgVal = orgVal.trim();
        if (hasOrgFilter && !filterOrg.equals(orgVal)) {
            continue;
        }

        // ---- 筛选：是否收款 ----
        if (hasReceiptFilter) {
            if ("1".equals(filterReceipt) && !isMarked) {
                continue;  // 筛选"是"时，隐藏待处理
            }
            if ("0".equals(filterReceipt) && isMarked) {
                continue;  // 筛选"否"时，隐藏已处理
            }
        }

        displayedCount++;
        if (isMarked) {
            markedCount++;
            displayedMarked++;
        }

        // 生产组织展示转换：0→杭州, 1→安吉
        String orgDisplay = orgVal;
        if ("0".equals(orgVal)) {
            orgDisplay = "杭州";
        } else if ("1".equals(orgVal)) {
            orgDisplay = "安吉";
        }

        // 生产订单状态展示转换：0-无,1-计划,2-计划确认,3-下达,4-开工,5-完工,6-结案,7-结算
        String[] statusMap = {"无","计划","计划确认","下达","开工","完工","结案","结算"};
        String statusDisplay = row.get("生产订单状态");
        String statusRaw = "";
        if (statusDisplay != null) {
            statusRaw = statusDisplay.trim();
            try {
                int st = Integer.parseInt(statusRaw);
                if (st >= 0 && st < statusMap.length) {
                    statusDisplay = st + "-" + statusMap[st];
                }
            } catch (NumberFormatException e) {}
        }

        String rowClass = isMarked ? "row-marked" : "";
        String chkDisabled = isMarked ? "disabled" : "";
        String chkId = "chk_" + fid;
%>
            <tr class="<%= rowClass %>" id="row_<%= fid %>">
                <td class="chk-col">
                    <% if (!isExport) { %>
                        <input type="checkbox" class="row-chk" id="<%= chkId %>"
                               value="<%= fid %>" <%= chkDisabled %>
                               onclick="onRowCheck(this)" />
                    <% } else if (isMarked) { %>
                        ✔
                    <% } %>
                </td>
                <% for (String col : displayCols) { %>
                    <%
                        String val = row.get(col);
                        if (val == null) val = "";
                        // 生产组织列：转换为中文
                        if (col.equals("生产组织")) {
                            val = orgDisplay;
                        }
                        // 生产订单状态列：转换为文本
                        if (col.equals("生产订单状态")) {
                            val = statusDisplay;
                        }
                        // 未回数量：保留整数
                        if (col.equals("未回数量") && !val.isEmpty()) {
                            try {
                                double d = Double.parseDouble(val);
                                val = String.valueOf((long) d);
                            } catch (NumberFormatException e) {}
                        }
                        // 数字列右对齐
                        boolean isNum = col.equals("未回数量");
                    %>
                    <td class="<%= isNum ? "num" : "" %>"><%= val %></td>
                <% } %>
                <td>
                    <% if (isMarked) { %>
                        <span class="marked-tag">已收款 <%= markedDate.get(fid) != null ? markedDate.get(fid) : "" %></span>
                    <% } else { %>
                        <span style="color:#e65100;">待处理</span>
                    <% } %>
                </td>
            </tr>
<%
    }
    int pendingCount = totalCount - markedCount;
    int displayedPending = displayedCount - displayedMarked;
%>
        </tbody>
    </table>

<% if (!isExport) { %>
<%
    // 构建带筛选条件的导出 URL
    StringBuilder exportUrlBuilder = new StringBuilder();
    exportUrlBuilder.append(request.getRequestURI()).append("?export=excel");
    if (hasFilter) {
        if (!filterDateFrom.isEmpty()) exportUrlBuilder.append("&dateFrom=").append(URLEncoder.encode(filterDateFrom, "UTF-8"));
        if (!filterDateTo.isEmpty()) exportUrlBuilder.append("&dateTo=").append(URLEncoder.encode(filterDateTo, "UTF-8"));
        if (!filterOrg.isEmpty()) exportUrlBuilder.append("&orgFilter=").append(URLEncoder.encode(filterOrg, "UTF-8"));
        if (!filterReceipt.isEmpty()) exportUrlBuilder.append("&receiptFilter=").append(URLEncoder.encode(filterReceipt, "UTF-8"));
    }
    String exportUrlWithFilter = exportUrlBuilder.toString();
%>
    <p style="margin-top:20px;color:#666;">
        显示记录数：<strong><%= displayedCount %></strong> / <%= totalCount %>
        &nbsp;&nbsp;（待处理: <span style="color:#e65100;font-weight:bold;"><%= displayedPending %></span>
        ，已收款: <span style="color:#2e7d32;font-weight:bold;"><%= displayedMarked %></span>）<%
        if (hasFilter) {
            out.print("&nbsp;&nbsp;<span style='color:#007acc;'>[筛选模式]</span>");
        }
        %>
        &nbsp;&nbsp;
        <a class="btn btn-export" href="<%= exportUrlWithFilter %>">导出Excel</a>
    </p>

<script>
    // ==================== 前端交互逻辑 ====================
    var checkedCount = 0;
    var totalPending = <%= displayedPending %>;

    // 统计栏
    document.getElementById('statPending').textContent = '<%= displayedPending %>';
    document.getElementById('statDone').textContent = '<%= displayedMarked %>';
    document.getElementById('statTotal').textContent = '<%= displayedCount %>';
    document.getElementById('statsBar').style.display = 'flex';

    // 全选/取消全选（仅可选的行）
    function toggleAll(masterChk) {
        var chks = document.getElementsByClassName('row-chk');
        checkedCount = 0;
        for (var i = 0; i < chks.length; i++) {
            if (!chks[i].disabled) {
                chks[i].checked = masterChk.checked;
                if (masterChk.checked) checkedCount++;
            }
        }
        updateBtn();
    }

    // 单个复选框变化
    function onRowCheck(chk) {
        if (chk.checked) {
            checkedCount++;
        } else {
            checkedCount--;
            // 取消全选状态
            document.getElementById('selectAll').checked = false;
        }
        updateBtn();
    }

    // 更新按钮状态
    function updateBtn() {
        var btn = document.getElementById('btnConfirm');
        btn.disabled = (checkedCount === 0);
        btn.textContent = checkedCount > 0 ? '✔ 确认收款（' + checkedCount + '条）' : '✔ 确认收款';
    }

    // 执行保存
    function doSave() {
        var chks = document.getElementsByClassName('row-chk');
        var selectedIds = [];
        for (var i = 0; i < chks.length; i++) {
            if (chks[i].checked && !chks[i].disabled) {
                selectedIds.push(chks[i].value);
            }
        }
        if (selectedIds.length === 0) {
            alert('请至少选择一行数据');
            return;
        }
        if (!confirm('确认将选中的 ' + selectedIds.length + ' 条记录标注为已收款吗？\n操作日期：<%= todayStr %>')) {
            return;
        }

        // 构造表单提交
        var container = document.getElementById('selectedIdsContainer');
        container.innerHTML = '';
        for (var i = 0; i < selectedIds.length; i++) {
            var input = document.createElement('input');
            input.type = 'hidden';
            input.name = 'selectedIds';
            input.value = selectedIds[i];
            container.appendChild(input);
        }

        var btn = document.getElementById('btnConfirm');
        btn.disabled = true;
        btn.textContent = '处理中...';

        document.getElementById('mainForm').submit();
    }

    // 页面加载后恢复按钮状态
    (function() {
        var chks = document.getElementsByClassName('row-chk');
        checkedCount = 0;
        for (var i = 0; i < chks.length; i++) {
            if (chks[i].checked) checkedCount++;
        }
        updateBtn();
    })();
</script>

</body>
</html>
<% } %>
