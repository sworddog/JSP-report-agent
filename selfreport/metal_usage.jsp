<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.sql.*,java.text.*" %>
<%
    request.setCharacterEncoding("UTF-8");

    String searchTerm = request.getParameter("spec");
    if (searchTerm == null) searchTerm = "";
    searchTerm = searchTerm.trim();
    boolean searched = !searchTerm.isEmpty();

    String selectedCode = "";
    String selectedName = "";
    boolean found = false;
    String errorMsg = "";

    // ---- 金属定义 ----
    String[] metalCodes = {"60100001", "60200001", "60300001", "60500001"};
    String[] metalNames = {"铜", "锡", "铝", "镍"};
    // 基准价（元/吨）
    double[] basePrice = {76800, 256900, 19760, 133500};
    // 现价（元/克），从采购订单查询 FTAXPRICE；兜底 = 基准价 / 1,000,000
    double[] currentPrice = {0.0768, 0.2569, 0.01976, 0.1335};

    String[] cats = {"箱体", "蜗轮", "蜗杆", "法兰", "端盖"};
    double[][] usage = new double[5][4];  // [类别][金属] 用量

    if (searched) {
        Connection conn = null;
        PreparedStatement pstmt = null;
        ResultSet rs = null;

        try {
            Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
            String url = "jdbc:sqlserver://172.16.5.185:1433;databaseName=AISmonth6cs";
            conn = DriverManager.getConnection(url, "onlyreaduser", "Supror@2003");

            // ---- Step 1: 查询现价（采购订单最新含税单价） ----
            String priceSQL =
                "WITH LATEST_PRICE AS ( " +
                "  SELECT wl.FNUMBER, c.FTAXPRICE, a.FDATE, " +
                "    ROW_NUMBER() OVER (PARTITION BY wl.FNUMBER ORDER BY a.FDATE DESC, a.FID DESC) AS RN " +
                "  FROM T_PUR_POORDER a " +
                "  INNER JOIN T_PUR_POORDERENTRY b ON b.FID = a.FID " +
                "  INNER JOIN T_PUR_POORDERENTRY_F c ON c.FENTRYID = b.FENTRYID " +
                "  INNER JOIN T_BD_MATERIAL wl ON wl.FMATERIALID = b.FMATERIALID " +
                "  WHERE wl.FNUMBER IN ('60100001','60200001','60300001','60500001') " +
                ") " +
                "SELECT FNUMBER, FTAXPRICE FROM LATEST_PRICE WHERE RN = 1";

            pstmt = conn.prepareStatement(priceSQL);
            rs = pstmt.executeQuery();
            while (rs.next()) {
                String code = rs.getString("FNUMBER");
                double price = rs.getDouble("FTAXPRICE");
                for (int i = 0; i < 4; i++) {
                    if (metalCodes[i].equals(code)) {
                        currentPrice[i] = price;
                        break;
                    }
                }
            }
            rs.close();
            pstmt.close();

            // ---- Step 2: 模糊搜索成品物料 ----
            String searchSQL =
                "SELECT TOP 1 m.FNUMBER, ml.FNAME " +
                "FROM T_BD_MATERIAL m " +
                "INNER JOIN T_BD_MATERIAL_L ml ON ml.FMATERIALID = m.FMATERIALID " +
                "INNER JOIN T_ENG_BOM b ON b.FMATERIALID = m.FMATERIALID " +
                "WHERE m.FNUMBER LIKE '1%' AND ml.FNAME LIKE ? " +
                "ORDER BY m.FNUMBER ASC";

            pstmt = conn.prepareStatement(searchSQL);
            pstmt.setString(1, searchTerm + "%");
            rs = pstmt.executeQuery();

            if (rs.next()) {
                selectedCode = rs.getString("FNUMBER");
                selectedName = rs.getString("FNAME");
                found = true;
            }
            rs.close();
            pstmt.close();

            // ---- Step 3: 查询五大子项金属用量 ----
            if (found) {
                String metalSQL =
                    "WITH CATEGORIES AS (" +
                    "  SELECT '201%' AS PREFIX, N'箱体' AS CAT, 1 AS SORT " +
                    "  UNION ALL SELECT '202%', N'蜗轮', 2 " +
                    "  UNION ALL SELECT '203%', N'蜗杆', 3 " +
                    "  UNION ALL SELECT '204%', N'法兰', 4 " +
                    "  UNION ALL SELECT '206%', N'端盖', 5 " +
                    "), " +
                    "TARGET AS ( " +
                    "  SELECT FMATERIALID FROM T_BD_MATERIAL WHERE FNUMBER = ? " +
                    "), " +
                    "BOM_LATEST AS ( " +
                    "  SELECT FMATERIALID, FID FROM ( " +
                    "    SELECT FMATERIALID, FID, " +
                    "      ROW_NUMBER() OVER (PARTITION BY FMATERIALID ORDER BY FNUMBER DESC) AS RN " +
                    "    FROM T_ENG_BOM WHERE FUSEORGID = 1 " +
                    "  ) x WHERE RN = 1 " +
                    "), " +
                    "BOM_FULL AS ( " +
                    "  SELECT " +
                    "    CASE WHEN m.FNUMBER LIKE '201%' THEN N'箱体' " +
                    "         WHEN m.FNUMBER LIKE '202%' THEN N'蜗轮' " +
                    "         WHEN m.FNUMBER LIKE '203%' THEN N'蜗杆' " +
                    "         WHEN m.FNUMBER LIKE '204%' THEN N'法兰' " +
                    "         WHEN m.FNUMBER LIKE '206%' THEN N'端盖' " +
                    "    END AS CAT, " +
                    "    bc.FMATERIALID AS CHILD_ID, " +
                    "    CAST(bc.FNUMERATOR * 1.0 / NULLIF(bc.FDENOMINATOR, 0) AS DECIMAL(18,6)) AS RATIO, " +
                    "    1 AS LVL " +
                    "  FROM TARGET t " +
                    "  INNER JOIN BOM_LATEST b ON b.FMATERIALID = t.FMATERIALID " +
                    "  INNER JOIN T_ENG_BOMCHILD bc ON bc.FID = b.FID " +
                    "  INNER JOIN T_BD_MATERIAL m ON m.FMATERIALID = bc.FMATERIALID " +
                    "  UNION ALL " +
                    "  SELECT " +
                    "    ISNULL(bf.CAT, " +
                    "      CASE WHEN m2.FNUMBER LIKE '201%' THEN N'箱体' " +
                    "           WHEN m2.FNUMBER LIKE '202%' THEN N'蜗轮' " +
                    "           WHEN m2.FNUMBER LIKE '203%' THEN N'蜗杆' " +
                    "           WHEN m2.FNUMBER LIKE '204%' THEN N'法兰' " +
                    "           WHEN m2.FNUMBER LIKE '206%' THEN N'端盖' " +
                    "      END), " +
                    "    bc2.FMATERIALID, " +
                    "    CAST(bf.RATIO * (bc2.FNUMERATOR * 1.0 / NULLIF(bc2.FDENOMINATOR, 0)) AS DECIMAL(18,6)) AS RATIO, " +
                    "    bf.LVL + 1 " +
                    "  FROM BOM_FULL bf " +
                    "  INNER JOIN BOM_LATEST b2 ON b2.FMATERIALID = bf.CHILD_ID " +
                    "  INNER JOIN T_ENG_BOMCHILD bc2 ON bc2.FID = b2.FID " +
                    "  INNER JOIN T_BD_MATERIAL m2 ON m2.FMATERIALID = bc2.FMATERIALID " +
                    "  WHERE bf.LVL < 6 " +
                    ") " +
                    "SELECT cat.CAT, cat.SORT, " +
                    "  ISNULL(metal.CU, 0) AS CU, " +
                    "  ISNULL(metal.SN, 0) AS SN, " +
                    "  ISNULL(metal.AL, 0) AS AL, " +
                    "  ISNULL(metal.NI, 0) AS NI " +
                    "FROM CATEGORIES cat " +
                    "LEFT JOIN ( " +
                    "  SELECT bf.CAT, " +
                    "    SUM(CASE WHEN m.FNUMBER = '60100001' THEN bf.RATIO ELSE 0 END) AS CU, " +
                    "    SUM(CASE WHEN m.FNUMBER = '60200001' THEN bf.RATIO ELSE 0 END) AS SN, " +
                    "    SUM(CASE WHEN m.FNUMBER = '60300001' THEN bf.RATIO ELSE 0 END) AS AL, " +
                    "    SUM(CASE WHEN m.FNUMBER = '60500001' THEN bf.RATIO ELSE 0 END) AS NI " +
                    "  FROM BOM_FULL bf " +
                    "  INNER JOIN T_BD_MATERIAL m ON m.FMATERIALID = bf.CHILD_ID " +
                    "  WHERE m.FNUMBER IN ('60100001','60200001','60300001','60500001') " +
                    "    AND bf.CAT IS NOT NULL " +
                    "  GROUP BY bf.CAT " +
                    ") metal ON metal.CAT = cat.CAT " +
                    "ORDER BY cat.SORT";

                pstmt = conn.prepareStatement(metalSQL);
                pstmt.setString(1, selectedCode);
                rs = pstmt.executeQuery();

                while (rs.next()) {
                    int sort = rs.getInt("SORT") - 1;
                    if (sort >= 0 && sort < 5) {
                        usage[sort][0] = rs.getDouble("CU");
                        usage[sort][1] = rs.getDouble("SN");
                        usage[sort][2] = rs.getDouble("AL");
                        usage[sort][3] = rs.getDouble("NI");
                    }
                }
            }
        } catch (Exception e) {
            errorMsg = "查询出错：" + e.getMessage();
        } finally {
            try { if (rs != null) rs.close(); } catch (Exception e) {}
            try { if (pstmt != null) pstmt.close(); } catch (Exception e) {}
            try { if (conn != null) conn.close(); } catch (Exception e) {}
        }
    }
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>单一产品金属原材料用料查询</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px 30px; }
        h1 { color: #333; font-size: 20px; margin-bottom: 16px; }
        .search-bar { display: flex; align-items: center; gap: 12px; flex-wrap: wrap;
            margin-bottom: 20px; padding: 12px 16px; background: #f9f9f9;
            border: 1px solid #e0e0e0; border-radius: 4px; }
        .search-bar label { font-size: 14px; }
        .search-bar input[type="text"] { padding: 6px 10px; border: 1px solid #ccc;
            border-radius: 3px; font-size: 14px; width: 300px; }
        .btn { padding: 7px 18px; border: none; border-radius: 4px; cursor: pointer;
            font-size: 14px; font-weight: bold; color: #fff; }
        .btn-search { background: #007acc; }
        .btn-search:hover { background: #005fa3; }
        .msg { padding: 8px 16px; border-radius: 4px; margin-bottom: 12px; font-size: 14px; }
        .msg-error { background: #ffebee; color: #c62828; border: 1px solid #ef9a9a; }
        .msg-info { background: #e3f2fd; color: #1565c0; border: 1px solid #90caf9; }
        .msg-warn { background: #fff3e0; color: #e65100; border: 1px solid #ffcc80; }
        .result-header { margin-bottom: 8px; font-size: 15px; }
        .result-header b { color: #007acc; }

        /* ---- 表格样式 ---- */
        .table-wrap { overflow-x: auto; max-width: 1000px; margin-top: 10px; }
        table { border-collapse: collapse; width: 100%; font-size: 13px; }
        th, td { border: 1px solid #aaa; padding: 7px 10px; text-align: center; }
        th { background-color: #e8e8e8; white-space: nowrap; }
        th.cat-col { text-align: left; width: 60px; }
        td.cat-col { text-align: left; font-weight: bold; background-color: #f5f5f5; }
        /* 金属分组颜色 */
        th.gr-cu { background-color: #ffcc80; }
        th.gr-sn { background-color: #b3e5fc; }
        th.gr-al { background-color: #c8e6c9; }
        th.gr-ni { background-color: #e1bee7; }
        td.num { text-align: right; font-family: Consolas, monospace; }
        td.zero { color: #bbb; }
        td.has-val { color: #333; }
        td.diff-up { color: #c62828; font-weight: bold; }
        /* 底部现价栏 */
        tr.price-footer td { background-color: #fffde7; font-size: 13px; text-align: left; padding: 10px 12px; }
        tr.price-footer td b { color: #c62828; }
        tr.total-row td { background-color: #fafafa; font-weight: bold; }
        /* 表头分组线 */
        th.group-cu { border-bottom: 3px solid #e65100; }
        th.group-sn { border-bottom: 3px solid #01579b; }
        th.group-al { border-bottom: 3px solid #1b5e20; }
        th.group-ni { border-bottom: 3px solid #6a1b9a; }

        .legend { margin-top: 6px; font-size: 12px; color: #888; }
    </style>
</head>
<body>

<h1>产品主要金属原材料成本查询</h1>

<div class="search-bar">
    <form method="get" action="" style="display: flex; align-items: center; gap: 12px; flex-wrap: wrap;">
        <label>成品规格型号：
            <input type="text" name="spec" value="<%= searchTerm %>"
                placeholder="如 WPDS80-10-0.75" />
        </label>
        <button type="submit" class="btn btn-search">查询</button>
        <% if (searched) { %>
            <a href="<%= request.getRequestURI() %>" style="padding: 5px 14px; background: #999; color: #fff;
                text-decoration: none; border-radius: 3px; font-size: 14px;">清除</a>
        <% } %>
    </form>
</div>

<% if (!errorMsg.isEmpty()) { %>
    <div class="msg msg-error"><%= errorMsg %></div>
<% } %>

<% if (searched) { %>
    <% if (!found) { %>
        <div class="msg msg-warn">
            未找到与 "<b><%= searchTerm %></b>" 匹配的成品物料（编码1开头且有BOM）。
        </div>
    <% } else { %>
        <div class="msg msg-info result-header">
            ✅ 匹配物料：<b><%= selectedCode %></b> &nbsp; <%= selectedName %>
        </div>

        <%-- ========== 表一：用量表 ========== --%>
        <h2 style="font-size:16px; margin: 20px 0 8px 0; color:#555;">📋 金属用量（单位：克）</h2>
        <div class="table-wrap">
        <table>
            <thead>
                <tr>
                    <th class="cat-col">类别</th>
                    <th class="gr-cu">铜用量</th>
                    <th class="gr-sn">锡用量</th>
                    <th class="gr-al">铝用量</th>
                    <th class="gr-ni">镍用量</th>
                </tr>
            </thead>
            <tbody>
                <%
                double[] totalUsage = new double[4];
                for (int i = 0; i < 5; i++) {
                %>
                <tr>
                    <td class="cat-col"><%= cats[i] %></td>
                    <% for (int j = 0; j < 4; j++) {
                        totalUsage[j] += usage[i][j];
                    %>
                    <td class="num <%= usage[i][j] > 0 ? "has-val" : "zero" %>">
                        <%= String.format("%.2f", usage[i][j]) %>
                    </td>
                    <% } %>
                </tr>
                <% } %>
                <tr class="total-row">
                    <td class="cat-col">合计</td>
                    <% for (int j = 0; j < 4; j++) { %>
                    <td class="num"><%= String.format("%.2f", totalUsage[j]) %></td>
                    <% } %>
                </tr>
            </tbody>
        </table>
        </div>

        <%-- ========== 表二：价格表 ========== --%>
        <h2 style="font-size:16px; margin: 24px 0 8px 0; color:#555;">💰 金属价格（单位：元）</h2>
        <div class="table-wrap">
        <table>
            <thead>
                <tr>
                    <th class="cat-col" rowspan="2">类别</th>
                    <th class="group-cu" colspan="3">铜</th>
                    <th class="group-sn" colspan="3">锡</th>
                    <th class="group-al" colspan="3">铝</th>
                    <th class="group-ni" colspan="3">镍</th>
                </tr>
                <tr>
                    <th class="gr-cu">基准价</th><th class="gr-cu">现价</th><th class="gr-cu">成本差异</th>
                    <th class="gr-sn">基准价</th><th class="gr-sn">现价</th><th class="gr-sn">成本差异</th>
                    <th class="gr-al">基准价</th><th class="gr-al">现价</th><th class="gr-al">成本差异</th>
                    <th class="gr-ni">基准价</th><th class="gr-ni">现价</th><th class="gr-ni">成本差异</th>
                </tr>
            </thead>
            <tbody>
                <%
                double[] totalBase = new double[4];
                double[] totalCurrent = new double[4];
                double[] totalDiff = new double[4];
                for (int i = 0; i < 5; i++) {
                %>
                <tr>
                    <td class="cat-col"><%= cats[i] %></td>
                    <% for (int j = 0; j < 4; j++) {
                        double baseAmt = usage[i][j] * basePrice[j] / 1000000;  // 克 × 元/吨 ÷ 1,000,000
                        double currAmt = usage[i][j] * currentPrice[j];          // 克 × 元/克
                        double diff    = currAmt - baseAmt;                       // 成本差异 = 现价 - 基准价
                        totalBase[j]    += baseAmt;
                        totalCurrent[j] += currAmt;
                        totalDiff[j]    += diff;
                        boolean hasUsage = usage[i][j] > 0;
                    %>
                    <td class="num <%= hasUsage ? "has-val" : "zero" %>">
                        <%= String.format("%.2f", baseAmt) %>
                    </td>
                    <td class="num <%= hasUsage ? "has-val" : "zero" %>">
                        <%= String.format("%.2f", currAmt) %>
                    </td>
                    <td class="num <%= hasUsage ? (diff > 0 ? "diff-up" : "has-val") : "zero" %>">
                        <%= String.format("%.2f", diff) %>
                    </td>
                    <% } %>
                </tr>
                <% } %>
                <tr class="total-row">
                    <td class="cat-col">合计</td>
                    <% for (int j = 0; j < 4; j++) { %>
                    <td class="num"><%= String.format("%.2f", totalBase[j]) %></td>
                    <td class="num"><%= String.format("%.2f", totalCurrent[j]) %></td>
                    <td class="num <%= totalDiff[j] > 0 ? "diff-up" : "" %>">
                        <%= String.format("%.2f", totalDiff[j]) %>
                    </td>
                    <% } %>
                </tr>
            </tbody>
        </table>
        </div>

        <%-- 底部：单价说明 --%>
        <%
            StringBuilder priceInfo = new StringBuilder();
            for (int j = 0; j < 4; j++) {
                double pricePerTon = currentPrice[j] * 1000000; // 元/克 → 元/吨
                priceInfo.append(metalNames[j]).append("：")
                         .append(String.format("%.0f", pricePerTon)).append(" 元/吨");
                if (j < 3) priceInfo.append("；");
            }
        %>
        <div class="msg msg-info" style="margin-top:16px; font-size:13px;">
            📊 金属现价&nbsp;&nbsp;
            <%= priceInfo.toString() %>
            &nbsp;&nbsp;|&nbsp;&nbsp;
            基准单价（元/吨）：铜 76800元/吨 ； 锡 256900元/吨 ； 铝 19760元/吨 ； 镍 133500元/吨
        </div>

        <p class="legend">注：现价单位元/克（取自 <code>最新采购订单单价</code>）；价格 = 用量 × 单价。</p>
    <% } %>
<% } else { %>
    <div class="msg msg-info">
        💡 请输入成品规格型号（支持模糊搜索），系统将自动匹配一个最接近的成品物料，展示其五大子项的金属用量及价格。
    </div>
<% } %>

</body>
</html>
