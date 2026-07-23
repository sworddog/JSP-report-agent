<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.sql.*,java.util.*,java.net.*" %>
<%
    // ==================== 导出模式判断 ====================
    boolean isExport = "excel".equals(request.getParameter("export"));
    if (isExport) {
        response.setContentType("application/vnd.ms-excel; charset=UTF-8");
        response.setHeader("Content-Disposition", "attachment; filename=" +
            URLEncoder.encode("蛟龙颜旭发料统计报表.xls", "UTF-8"));
    }
%>
<% if (!isExport) { %>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>蛟龙 / 颜旭 发料统计报表</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; }
        h1 { color: #333; }
        table { border-collapse: collapse; width: 95%; margin-top: 20px; font-size: 14px; }
        th, td { border: 1px solid #ccc; padding: 6px 10px; text-align: left; }
        th { background-color: #f5f5f5; white-space: nowrap; }
        .num { text-align: right; }
        .error { color: red; }
    </style>
</head>
<body>
    <h1>蛟龙 / 颜旭 发料统计报表</h1>
<% } %>

<%
    // ==================== 获取查询参数 ====================
    request.setCharacterEncoding("UTF-8");
    String filterBillNo = request.getParameter("billNo");
    if (filterBillNo == null) filterBillNo = "";
    filterBillNo = filterBillNo.trim();

    String filterSupplier = request.getParameter("supplier");
    if (filterSupplier == null) filterSupplier = "";
    filterSupplier = filterSupplier.trim();

    boolean hasFilter = !filterBillNo.isEmpty() || !filterSupplier.isEmpty();

    // 构建导出 URL（保留当前筛选条件）
    String exportUrl = request.getRequestURI() + "?export=excel";
    if (hasFilter) {
        exportUrl += "&billNo=" + URLEncoder.encode(filterBillNo, "UTF-8")
                  + "&supplier=" + URLEncoder.encode(filterSupplier, "UTF-8");
    }
%>

    <%-- 查询过滤表单 --%>
    <% if (!isExport) { %>
    <div style="margin-bottom: 16px; padding: 12px; background: #f9f9f9; border: 1px solid #e0e0e0; border-radius: 4px;">
        <form method="get" action="" style="display: flex; align-items: center; gap: 12px; flex-wrap: wrap;">
            <label>单据编号：<input type="text" name="billNo" value="<%= filterBillNo %>" placeholder="输入单据编号" style="padding: 4px 8px; border: 1px solid #ccc; border-radius: 3px; width: 180px;" /></label>
            <label>供应商：<input type="text" name="supplier" value="<%= filterSupplier %>" placeholder="输入供应商" style="padding: 4px 8px; border: 1px solid #ccc; border-radius: 3px; width: 180px;" /></label>
            <button type="submit" style="padding: 5px 14px; background: #007acc; color: #fff; border: none; border-radius: 3px; cursor: pointer;">查询</button>
            <% if (hasFilter) { %>
                <a href="<%= request.getRequestURI() %>" style="padding: 5px 14px; background: #999; color: #fff; text-decoration: none; border-radius: 3px;">清除</a>
            <% } %>
            <span style="flex-grow: 1;"></span>
            <a href="<%= exportUrl %>" style="padding: 5px 16px; background: #28a745; color: #fff; text-decoration: none; border-radius: 3px; font-weight: bold;">导出Excel</a>
        </form>
    </div>
    <% } %>

<%
    // ==================== 查询第一个数据库（ecology） ====================
    Connection conn1 = null;
    Statement stmt1 = null;
    ResultSet rs1 = null;

    Map<Integer, Map<String, String>> data1 = new LinkedHashMap<Integer, Map<String, String>>();

    try {
        Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");

        String url1 = "jdbc:sqlserver://172.16.5.188:1433;databaseName=ecology";
        conn1 = DriverManager.getConnection(url1, "sa", "Sble123456");
        stmt1 = conn1.createStatement();
        String sql1 = "SELECT [fentryid],[蛟龙累计发料数量],[颜旭累计发料数量] FROM view_jlfhjl";
        rs1 = stmt1.executeQuery(sql1);

        while (rs1.next()) {
            Map<String, String> row = new HashMap<String, String>();
            int id = rs1.getInt("fentryid");
            row.put("蛟龙累计发料数量", rs1.getString("蛟龙累计发料数量"));
            row.put("颜旭累计发料数量", rs1.getString("颜旭累计发料数量"));
            data1.put(id, row);
        }
    } catch (Exception e) {
        out.println("<p class='error'>数据库1（ecology）连接或查询出错：" + e.getMessage() + "</p>");
    } finally {
        try { if (rs1 != null) rs1.close(); } catch (Exception e) {}
        try { if (stmt1 != null) stmt1.close(); } catch (Exception e) {}
        try { if (conn1 != null) conn1.close(); } catch (Exception e) {}
    }

    // ==================== 查询第二个数据库（AISmonth6cs，作为主表） ====================
    Connection conn2 = null;
    Statement stmt2 = null;
    ResultSet rs2 = null;

    Map<Integer, Map<String, String>> data2 = new LinkedHashMap<Integer, Map<String, String>>();

    try {
        Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");

        String url2 = "jdbc:sqlserver://172.16.5.185:1433;databaseName=AISmonth6cs";
        conn2 = DriverManager.getConnection(url2, "onlyreaduser", "Supror@2003");
        stmt2 = conn2.createStatement();
        String sql2 = "SELECT [FENTRYID],[单据编号],[采购日期],[供应商],[采购订单行号],[规格型号],[采购数量],[交货日期],[备注],[剩余未发数量],[物料编码] " +
                      "FROM VIEW_TO_OA_CGDD ORDER BY [采购日期]";
        rs2 = stmt2.executeQuery(sql2);

        while (rs2.next()) {
            Map<String, String> row = new HashMap<String, String>();
            int id = rs2.getInt("FENTRYID");
            row.put("单据编号", rs2.getString("单据编号"));
            row.put("采购日期", rs2.getString("采购日期"));
            row.put("供应商", rs2.getString("供应商"));
            row.put("采购订单行号", rs2.getString("采购订单行号"));
            row.put("规格型号", rs2.getString("规格型号"));
            row.put("采购数量", rs2.getString("采购数量"));
            row.put("交货日期", rs2.getString("交货日期"));
            row.put("备注", rs2.getString("备注"));
            row.put("剩余未发数量", rs2.getString("剩余未发数量"));
            row.put("物料编码", rs2.getString("物料编码"));
            data2.put(id, row);
        }
    } catch (Exception e) {
        out.println("<p class='error'>数据库2（AISmonth6cs）连接或查询出错：" + e.getMessage() + "</p>");
    } finally {
        try { if (rs2 != null) rs2.close(); } catch (Exception e) {}
        try { if (stmt2 != null) stmt2.close(); } catch (Exception e) {}
        try { if (conn2 != null) conn2.close(); } catch (Exception e) {}
    }
%>

    <table>
        <thead>
            <tr>
                <th>单据编号</th>
                <th>采购日期</th>
                <th>供应商</th>
                <th>规格型号</th>
                <th>采购数量</th>
                <th>剩余未发数量</th>
                <th>交货日期</th>
                <th>采购订单行号</th>
                <th>蛟龙累计发料数量</th>
                <th>蛟龙未发数量</th>
                <th>颜旭累计发料数量</th>
                <th>颜旭未发数量</th>
                <th>备注</th>
            </tr>
        </thead>
        <tbody>
<%
    int displayedCount = 0;
    for (Map.Entry<Integer, Map<String, String>> entry : data2.entrySet()) {
        int id = entry.getKey();
        Map<String, String> row2 = entry.getValue();
        Map<String, String> row1 = data1.get(id);

        // 过滤：单据编号
        String billNo = row2.get("单据编号");
        if (!filterBillNo.isEmpty() && (billNo == null || !billNo.contains(filterBillNo))) {
            continue;
        }
        // 过滤：供应商
        String supplier = row2.get("供应商");
        if (!filterSupplier.isEmpty() && (supplier == null || !supplier.contains(filterSupplier))) {
            continue;
        }
        displayedCount++;

        String purchaseQtyStr = row2.get("采购数量");
        double purchaseQty = 0;
        if (purchaseQtyStr != null && !purchaseQtyStr.trim().isEmpty()) {
            try { purchaseQty = Double.parseDouble(purchaseQtyStr); } catch (NumberFormatException e) {}
        }

        String jlQtyStr = (row1 != null) ? row1.get("蛟龙累计发料数量") : null;
        double jlQty = 0;
        if (jlQtyStr != null && !jlQtyStr.trim().isEmpty()) {
            try { jlQty = Double.parseDouble(jlQtyStr); } catch (NumberFormatException e) {}
        }

        String yxQtyStr = (row1 != null) ? row1.get("颜旭累计发料数量") : null;
        double yxQty = 0;
        if (yxQtyStr != null && !yxQtyStr.trim().isEmpty()) {
            try { yxQty = Double.parseDouble(yxQtyStr); } catch (NumberFormatException e) {}
        }

        // 从视图直接获取剩余未发数量
        String remainQtyStr = row2.get("剩余未发数量");
        double remainQty = 0;
        if (remainQtyStr != null && !remainQtyStr.trim().isEmpty()) {
            try { remainQty = Double.parseDouble(remainQtyStr); } catch (NumberFormatException e) {}
        }

        double jlUnsent = purchaseQty - jlQty;
        double yxUnsent = purchaseQty - yxQty;

        // 物料编码3开头时，颜旭未发数量固定为空
        String materialCode = row2.get("物料编码");
        boolean isMat3 = (materialCode != null && materialCode.startsWith("3"));
%>
            <tr>
                <td><%= row2.get("单据编号") != null ? row2.get("单据编号") : "" %></td>
                <td><%= row2.get("采购日期") != null ? row2.get("采购日期") : "" %></td>
                <td><%= row2.get("供应商") != null ? row2.get("供应商") : "" %></td>
                <td><%= row2.get("规格型号") != null ? row2.get("规格型号") : "" %></td>
                <td class="num"><%= purchaseQtyStr != null ? purchaseQtyStr : "" %></td>
                <td class="num"><%= remainQtyStr != null ? remainQtyStr : "" %></td>
                <td><%= row2.get("交货日期") != null ? row2.get("交货日期") : "" %></td>
                <td><%= row2.get("采购订单行号") != null ? row2.get("采购订单行号") : "" %></td>
                <td class="num"><%= jlQtyStr != null ? jlQtyStr : "" %></td>
                <td class="num"><%= jlUnsent %></td>
                <td class="num"><%= yxQtyStr != null ? yxQtyStr : "" %></td>
                <td class="num"><%= isMat3 ? "" : yxUnsent %></td>
                <td><%= row2.get("备注") != null ? row2.get("备注") : "" %></td>
            </tr>
<%
    }
%>
        </tbody>
    </table>

    <% if (!isExport) { %>
    <p style="margin-top:20px;color:#666;">
        <%
            int totalErp = data2.size();
            int oaMatch = countIntersection(data1, data2);
        %>
        显示记录数：<strong><%= displayedCount %></strong> / <%= totalErp %>
        &nbsp;&nbsp;（ERP: <%= totalErp %> 条，OA 匹配: <%= oaMatch %> 条）<%
        if (hasFilter) {
            out.print("&nbsp;&nbsp;<span style='color:#007acc;'>[筛选模式]</span>");
        }
        %>
        &nbsp;&nbsp;
        <a href="<%= exportUrl %>" style="padding: 5px 16px; background: #28a745; color: #fff; text-decoration: none; border-radius: 3px; font-weight: bold;">导出Excel</a>
    </p>

</body>
</html>
<% } %>

<%!
    private int countIntersection(Map<Integer, Map<String, String>> m1, Map<Integer, Map<String, String>> m2) {
        int count = 0;
        for (Integer key : m1.keySet()) {
            if (m2.containsKey(key)) count++;
        }
        return count;
    }
%>
