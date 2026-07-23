<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.sql.*,java.util.*,java.net.*,java.text.*" %>
<%
    // ==================== 导出模式判断 ====================
    boolean isExport = "excel".equals(request.getParameter("export"));
    if (isExport) {
        response.setContentType("application/vnd.ms-excel; charset=UTF-8");
        response.setHeader("Content-Disposition", "attachment; filename=" +
            URLEncoder.encode("外贸客户销售产品价分析表.xls", "UTF-8"));
    }
%>
<% if (!isExport) { %>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>外贸客户销售产品价分析表</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px 30px; }
        h1 { color: #333; font-size: 20px; }
        .filter-bar { display: flex; align-items: center; gap: 12px; flex-wrap: wrap;
            margin-bottom: 16px; padding: 12px 16px; background: #f9f9f9;
            border: 1px solid #e0e0e0; border-radius: 4px; }
        .filter-bar label { font-size: 14px; }
        .filter-bar input[type="text"], .filter-bar input[type="date"] {
            padding: 6px 10px; border: 1px solid #ccc; border-radius: 3px; font-size: 14px; }
        .btn { padding: 7px 18px; border: none; border-radius: 4px; cursor: pointer;
            font-size: 14px; font-weight: bold; color: #fff; text-decoration: none; display: inline-block; }
        .btn-search { background: #007acc; }
        .btn-search:hover { background: #005fa3; }
        .btn-export { background: #28a745; }
        .btn-export:hover { background: #1e7e34; }
        .btn-clear { background: #999; }
        .btn-clear:hover { background: #777; }
        .error { color: red; }
        .msg { padding: 8px 16px; border-radius: 4px; margin-bottom: 12px; font-size: 14px; }
        .msg-info { background: #e3f2fd; color: #1565c0; border: 1px solid #90caf9; }

        table { border-collapse: collapse; width: 100%; margin-top: 10px; font-size: 13px; }
        th, td { border: 1px solid #ccc; padding: 6px 10px; text-align: center; }
        th { background-color: #f5f5f5; white-space: nowrap; position: sticky; top: 0; }
        .num { text-align: right; font-family: Consolas, monospace; }
        .left { text-align: left; }
        th.col-cu { background-color: #ffcc80; }
        th.col-al { background-color: #c8e6c9; }
        th.col-sn { background-color: #b3e5fc; }
        th.col-ni { background-color: #e1bee7; }
        th.col-cost { background-color: #fff9c4; }
        th.col-base { background-color: #ffccbc; }
        th.col-qty { background-color: #d1c4e9; }
        .zero { color: #bbb; }
        .price-bar { margin-top: 12px; padding: 8px 14px; background: #fffde7;
            border: 1px solid #ffe082; border-radius: 4px; font-size: 13px; color: #555; }
    </style>
</head>
<body>
    <h1>外贸客户销售产品价分析表</h1>
<% } %>

<%
    // ==================== 获取查询参数 ====================
    request.setCharacterEncoding("UTF-8");

    // 是否点击了"查询"按钮（避免初次打开页面就加载全部数据）
    boolean searched = "1".equals(request.getParameter("search"));

    // 默认日期：当年1月1日 ~ 当日
    java.util.Calendar cal = java.util.Calendar.getInstance();
    int currentYear = cal.get(java.util.Calendar.YEAR);
    String defaultDateFrom = currentYear + "-01-01";
    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
    String defaultDateTo = sdf.format(new java.util.Date());

    String filterDateFrom = request.getParameter("dateFrom");
    if (filterDateFrom == null || (filterDateFrom = filterDateFrom.trim()).isEmpty()) {
        filterDateFrom = defaultDateFrom;
    }

    String filterDateTo = request.getParameter("dateTo");
    if (filterDateTo == null || (filterDateTo = filterDateTo.trim()).isEmpty()) {
        filterDateTo = defaultDateTo;
    }

    String filterCustomer = request.getParameter("customer");
    if (filterCustomer == null) filterCustomer = "";
    filterCustomer = filterCustomer.trim();

    boolean hasFilter = !filterDateFrom.isEmpty() || !filterDateTo.isEmpty() || !filterCustomer.isEmpty();

    // 构建导出 URL（保留当前筛选条件 + search 标记以触发查询）
    StringBuilder exportUrlBuilder = new StringBuilder();
    exportUrlBuilder.append(request.getRequestURI()).append("?export=excel&search=1");
    if (hasFilter) {
        if (!filterDateFrom.isEmpty()) exportUrlBuilder.append("&dateFrom=").append(URLEncoder.encode(filterDateFrom, "UTF-8"));
        if (!filterDateTo.isEmpty()) exportUrlBuilder.append("&dateTo=").append(URLEncoder.encode(filterDateTo, "UTF-8"));
        if (!filterCustomer.isEmpty()) exportUrlBuilder.append("&customer=").append(URLEncoder.encode(filterCustomer, "UTF-8"));
    }
    String exportUrl = exportUrlBuilder.toString();
%>

    <%-- 查询过滤表单 --%>
    <% if (!isExport) { %>
    <div class="filter-bar">
        <form method="get" action="" style="display: flex; align-items: center; gap: 12px; flex-wrap: wrap;">
            <label>起始日期：
                <input type="date" name="dateFrom" value="<%= filterDateFrom %>" style="width: 150px;" />
            </label>
            <label>结束日期：
                <input type="date" name="dateTo" value="<%= filterDateTo %>" style="width: 150px;" />
            </label>
            <label>客户：
                <input type="text" name="customer" value="<%= filterCustomer %>" placeholder="输入客户名称" style="width: 180px;" />
            </label>
            <button type="submit" class="btn btn-search" name="search" value="1">查询</button>
            <% if (hasFilter) { %>
                <a href="<%= request.getRequestURI() %>" class="btn btn-clear">清除</a>
            <% } %>
            <span style="flex-grow: 1;"></span>
            <a href="<%= exportUrl %>" class="btn btn-export">导出Excel</a>
        </form>
    </div>
    <% } %>

<%
    // ==================== 金属价格定义 ====================
    // 物料编码与金属对应关系
    String[] metalCodes  = {"60100001", "60200001", "60300001", "60500001"};
    String[] metalNames  = {"铜", "锡", "铝", "镍"};
    // 基准价（元/吨）
    double[] basePrice   = {76800, 256900, 19760, 133500};
    // 现价（元/克）—— 默认值=基准价/1,000,000，后续从采购订单查询覆盖
    double[] currentPrice = {0.0768, 0.2569, 0.01976, 0.1335};

    // 共享数据容器（声明在外部，供表格渲染访问）
    Map<String, Map<String, String>> modelMap = new LinkedHashMap<String, Map<String, String>>();
    Map<String, double[]> metalUsageMap = new LinkedHashMap<String, double[]>();
    List<Map<String, Object>> reportRows = new ArrayList<Map<String, Object>>();

    if (searched) {

    // ==================== Step 1: 查询金属最新采购单价 ====================
    // 复用 metal_usage.jsp 中的价格查询逻辑
    // 价格取自 ≤ 结束日期的最新采购订单含税单价
    {
        Connection conn1 = null;
        PreparedStatement pstmt1 = null;
        ResultSet rs1 = null;

        try {
            Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
            String url1 = "jdbc:sqlserver://172.16.5.185:1433;databaseName=AISmonth6cs";
            conn1 = DriverManager.getConnection(url1, "onlyreaduser", "Supror@2003");

            String priceSQL;
            if (!filterDateTo.isEmpty()) {
                priceSQL =
                    "WITH LATEST_PRICE AS ( " +
                    "  SELECT wl.FNUMBER, c.FTAXPRICE, a.FDATE, " +
                    "    ROW_NUMBER() OVER (PARTITION BY wl.FNUMBER ORDER BY a.FDATE DESC, a.FID DESC) AS RN " +
                    "  FROM T_PUR_POORDER a " +
                    "  INNER JOIN T_PUR_POORDERENTRY b ON b.FID = a.FID " +
                    "  INNER JOIN T_PUR_POORDERENTRY_F c ON c.FENTRYID = b.FENTRYID " +
                    "  INNER JOIN T_BD_MATERIAL wl ON wl.FMATERIALID = b.FMATERIALID " +
                    "  WHERE wl.FNUMBER IN ('60100001','60200001','60300001','60500001') " +
                    "    AND a.FDATE <= ? " +
                    ") " +
                    "SELECT FNUMBER, FTAXPRICE FROM LATEST_PRICE WHERE RN = 1";
            } else {
                priceSQL =
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
            }

            pstmt1 = conn1.prepareStatement(priceSQL);
            if (!filterDateTo.isEmpty()) {
                pstmt1.setString(1, filterDateTo);
            }
            rs1 = pstmt1.executeQuery();
            while (rs1.next()) {
                String code = rs1.getString("FNUMBER");
                double price = rs1.getDouble("FTAXPRICE");
                for (int i = 0; i < 4; i++) {
                    if (metalCodes[i].equals(code)) {
                        currentPrice[i] = price;
                        break;
                    }
                }
            }
        } catch (Exception e) {
            if (!isExport) {
                out.println("<p class='error'>金属价格查询出错：" + e.getMessage() + "</p>");
            }
        } finally {
            try { if (rs1 != null) rs1.close(); } catch (Exception e) {}
            try { if (pstmt1 != null) pstmt1.close(); } catch (Exception e) {}
            try { if (conn1 != null) conn1.close(); } catch (Exception e) {}
        }
    }

    // ==================== Step 2: 查询销售订单数据 ====================
    // 获取成品型号和销售数量，按物料编码汇总
    // 说明：
    //   - 表名基于金蝶 K3 Cloud 标准结构：T_SAL_ORDER(主表) / T_SAL_ORDERENTRY(分录)
    //   - 过滤规则：RV产品 FNUMBER LIKE '1101%'，SRV产品 FNUMBER LIKE '1105%'
    {
        Connection conn2 = null;
        PreparedStatement pstmt2 = null;
        ResultSet rs2 = null;

        try {
            Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
            String url2 = "jdbc:sqlserver://172.16.5.185:1433;databaseName=AISmonth6cs";
            conn2 = DriverManager.getConnection(url2, "onlyreaduser", "Supror@2003");

            // 附表2 表关联关系：
            // T1: T_SAL_ORDER (主表)           T1.FDATE = 筛选日期
            // T2: T_SAL_ORDERENTRY (明细)       T1.FID=T2.FID   T2.FQTY=销售数量  T2.F_VBIK_JSDJ=销售单价本位币
            // wl1: T_BD_MATERIAL (物料)         wl1.FMATERIALID=T2.FMATERIALID
            // wl2: T_BD_MATERIAL_L (物料语言)   wl2.FMATERIALID=T2.FMATERIALID  wl2.FSPECIFICATION=型号
            // kh1: T_BD_CUSTOMER (客户)         kh1.FCUSTID=T1.FCUSTID
            // kh2: T_BD_CUSTOMER_L (客户语言)   kh2.FCUSTID=T1.FCUSTID  kh2.FSHORTNAME=客户
            //
            // 过滤规则：RV=1101开头, SRV=1105开头
            // 分组规则：规格精确到速比（如 RV75-80-90B14 → RV75-80），同组内：
            //   - 金属用量取 FNUMBER 最小的那条的 BOM
            //   - 销售数量求和
            //   - 销售单价取最大值

            StringBuilder salesSQL = new StringBuilder();
            salesSQL.append("SELECT wl1.FNUMBER AS 物料编码, wl2.FSPECIFICATION AS 型号, ");
            salesSQL.append("  SUM(T2.FQTY) AS 销售数量, ");
            salesSQL.append("  MAX(T2.F_VBIK_JSDJ) AS 销售单价, ");
            salesSQL.append("  kh2.FSHORTNAME AS 客户, ");
            salesSQL.append("  MAX(T1.FDATE) AS 最新日期 ");
            salesSQL.append("FROM T_SAL_ORDER T1 ");
            salesSQL.append("INNER JOIN T_SAL_ORDERENTRY T2 ON T2.FID = T1.FID ");
            salesSQL.append("INNER JOIN T_BD_MATERIAL wl1 ON wl1.FMATERIALID = T2.FMATERIALID ");
            salesSQL.append("INNER JOIN T_BD_MATERIAL_L wl2 ON wl2.FMATERIALID = T2.FMATERIALID AND wl2.FLOCALEID = 2052 ");
            salesSQL.append("LEFT JOIN T_BD_CUSTOMER kh1 ON kh1.FCUSTID = T1.FCUSTID ");
            salesSQL.append("LEFT JOIN T_BD_CUSTOMER_L kh2 ON kh2.FCUSTID = kh1.FCUSTID AND kh2.FLOCALEID = 2052 ");
            salesSQL.append("WHERE (wl1.FNUMBER LIKE '1101%' OR wl1.FNUMBER LIKE '1105%') ");
            // 只显示规格左包含 RV/SRV 的（排除 TRV、TSRV 等）
            salesSQL.append("AND (wl2.FSPECIFICATION LIKE 'RV%' OR wl2.FSPECIFICATION LIKE 'SRV%') ");
            if (!filterDateFrom.isEmpty()) {
                salesSQL.append("AND T1.FDATE >= ? ");
            }
            if (!filterDateTo.isEmpty()) {
                salesSQL.append("AND T1.FDATE <= ? ");
            }
            if (!filterCustomer.isEmpty()) {
                salesSQL.append("AND kh2.FSHORTNAME LIKE ? ");
            }
            salesSQL.append("GROUP BY wl1.FNUMBER, wl2.FSPECIFICATION, kh2.FSHORTNAME ");
            salesSQL.append("ORDER BY wl1.FNUMBER ASC");

            pstmt2 = conn2.prepareStatement(salesSQL.toString());
            int paramIdx = 1;
            if (!filterDateFrom.isEmpty()) {
                pstmt2.setString(paramIdx++, filterDateFrom);
            }
            if (!filterDateTo.isEmpty()) {
                pstmt2.setString(paramIdx++, filterDateTo);
            }
            if (!filterCustomer.isEmpty()) {
                pstmt2.setString(paramIdx++, "%" + filterCustomer + "%");
            }

            // 先收集原始行，再按截断规格分组
            List<Map<String, String>> rawSalesRows = new ArrayList<Map<String, String>>();
            rs2 = pstmt2.executeQuery();
            while (rs2.next()) {
                Map<String, String> row = new LinkedHashMap<String, String>();
                row.put("物料编码", rs2.getString("物料编码") != null ? rs2.getString("物料编码") : "");
                row.put("型号", rs2.getString("型号") != null ? rs2.getString("型号") : "");
                row.put("销售数量", rs2.getString("销售数量") != null ? rs2.getString("销售数量") : "0");
                row.put("销售单价", rs2.getString("销售单价") != null ? rs2.getString("销售单价") : "0");
                row.put("客户", rs2.getString("客户") != null ? rs2.getString("客户") : "");
                rawSalesRows.add(row);
            }

            // 规格截断 + 分组：RV75-80-90B14 → RV75-80
            java.util.regex.Pattern baseModelPattern = java.util.regex.Pattern.compile("^(RV|SRV)\\d+-\\d+");

            // modelMap: key=截断后的基础型号(如"RV75-80"), value={物料编码, 型号, 销售数量, 销售单价, 客户}
            for (Map<String, String> row : rawSalesRows) {
                String spec = row.get("型号");
                String matCode = row.get("物料编码");
                String baseModel = spec;
                java.util.regex.Matcher m = baseModelPattern.matcher(spec);
                if (m.find()) {
                    baseModel = m.group();
                }

                double rowQty = 0, rowPrice = 0;
                try { rowQty = Double.parseDouble(row.get("销售数量")); } catch (NumberFormatException e) {}
                try { rowPrice = Double.parseDouble(row.get("销售单价")); } catch (NumberFormatException e) {}

                if (modelMap.containsKey(baseModel)) {
                    Map<String, String> exist = modelMap.get(baseModel);
                    // 数量累加
                    try {
                        double existQty = Double.parseDouble(exist.get("销售数量"));
                        exist.put("销售数量", String.valueOf(existQty + rowQty));
                    } catch (NumberFormatException e) {}
                    // 价格取最新（最大值）
                    try {
                        double existPrice = Double.parseDouble(exist.get("销售单价"));
                        if (rowPrice > existPrice) {
                            exist.put("销售单价", row.get("销售单价"));
                        }
                    } catch (NumberFormatException e) {}
                    // 物料编码取最小的（FNUMBER 最小 = 先创建的 = BOM 最全的）
                    String existCode = exist.get("物料编码");
                    if (matCode.compareTo(existCode) < 0) {
                        exist.put("物料编码", matCode);
                    }
                    // 合并客户
                    String existCust = exist.get("客户");
                    String newCust = row.get("客户");
                    if (newCust != null && !newCust.isEmpty() && !existCust.contains(newCust)) {
                        exist.put("客户", existCust + ";" + newCust);
                    }
                } else {
                    Map<String, String> grp = new LinkedHashMap<String, String>();
                    grp.put("物料编码", matCode);
                    grp.put("型号", baseModel);
                    grp.put("销售数量", row.get("销售数量"));
                    grp.put("销售单价", row.get("销售单价"));
                    grp.put("客户", row.get("客户") != null ? row.get("客户") : "");
                    modelMap.put(baseModel, grp);
                }
            }
        } catch (Exception e) {
            if (!isExport) {
                out.println("<p class='error'>销售订单查询出错：" + e.getMessage() + "</p>");
            }
        } finally {
            try { if (rs2 != null) rs2.close(); } catch (Exception e) {}
            try { if (pstmt2 != null) pstmt2.close(); } catch (Exception e) {}
            try { if (conn2 != null) conn2.close(); } catch (Exception e) {}
        }
    }

    // ==================== Step 3: 逐个型号查询金属用量（BOM 展开） ====================
    // 复用 metal_usage.jsp 的 BOM 递归 CTE 逻辑
    // 对每个成品物料编码，展开 BOM 汇总 4 种金属总用量（克）
    // 存储结果: 基础型号 → double[4]{铜, 锡, 铝, 镍}
    // 注：BOM 只查每组中的第一个物料编码（FNUMBER最小），同组共用

    for (Map.Entry<String, Map<String, String>> entry : modelMap.entrySet()) {
        String baseModel = entry.getKey();
        String bomMatCode = entry.getValue().get("物料编码");  // 该组用于查BOM的物料编码
        double[] totalUsage = new double[4];
        Connection conn3 = null;
        PreparedStatement pstmt3 = null;
        ResultSet rs3 = null;

        try {
            Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
            String url3 = "jdbc:sqlserver://172.16.5.185:1433;databaseName=AISmonth6cs";
            conn3 = DriverManager.getConnection(url3, "onlyreaduser", "Supror@2003");

            // 与 metal_usage.jsp 完全一致的 BOM CTE 查询（汇总版：直接求和，不按类别分组）
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
                "SELECT " +
                "  SUM(CASE WHEN m.FNUMBER = '60100001' THEN bf.RATIO ELSE 0 END) AS CU, " +
                "  SUM(CASE WHEN m.FNUMBER = '60200001' THEN bf.RATIO ELSE 0 END) AS SN, " +
                "  SUM(CASE WHEN m.FNUMBER = '60300001' THEN bf.RATIO ELSE 0 END) AS AL, " +
                "  SUM(CASE WHEN m.FNUMBER = '60500001' THEN bf.RATIO ELSE 0 END) AS NI " +
                "FROM BOM_FULL bf " +
                "INNER JOIN T_BD_MATERIAL m ON m.FMATERIALID = bf.CHILD_ID " +
                "WHERE m.FNUMBER IN ('60100001','60200001','60300001','60500001') " +
                "  AND bf.CAT IS NOT NULL";

            pstmt3 = conn3.prepareStatement(metalSQL);
            pstmt3.setString(1, bomMatCode);
            rs3 = pstmt3.executeQuery();

            if (rs3.next()) {
                totalUsage[0] = rs3.getDouble("CU");  // 铜
                totalUsage[1] = rs3.getDouble("SN");  // 锡
                totalUsage[2] = rs3.getDouble("AL");  // 铝
                totalUsage[3] = rs3.getDouble("NI");  // 镍
            }
        } catch (Exception e) {
            // BOM 查询失败时用量保持为 0（常见于该物料无 BOM，静默处理）
        } finally {
            try { if (rs3 != null) rs3.close(); } catch (Exception e) {}
            try { if (pstmt3 != null) pstmt3.close(); } catch (Exception e) {}
            try { if (conn3 != null) conn3.close(); } catch (Exception e) {}
        }

        metalUsageMap.put(baseModel, totalUsage);
    }

    // ==================== Step 4: 组装报表数据 ====================
    // 每行: 型号 | 铜用量 | 铝用量 | 锡用量 | 镍用量 | 成本价汇总 | 基准价 | 销售数量

    for (String baseModel : modelMap.keySet()) {
        Map<String, String> salesRow = modelMap.get(baseModel);
        double[] usage = metalUsageMap.get(baseModel);
        if (usage == null) {
            usage = new double[4];
        }

        Map<String, Object> row = new LinkedHashMap<String, Object>();
        row.put("基础型号", baseModel);
        row.put("BOM物料编码", salesRow.get("物料编码"));  // 实际查BOM用的物料编码
        row.put("型号", salesRow.get("型号"));

        // 金属用量（克）—— 注意数组顺序: [0]=铜, [1]=锡, [2]=铝, [3]=镍
        row.put("铜用量", usage[0]);
        row.put("锡用量", usage[1]);
        row.put("铝用量", usage[2]);
        row.put("镍用量", usage[3]);

        // 成本价汇总 = Σ 用量(克) × 金属现价(元/克)
        double costSummary = 0;
        for (int i = 0; i < 4; i++) {
            costSummary += usage[i] * currentPrice[i];
        }
        row.put("成本价汇总", costSummary);

        // 基准价 = 销售订单本位币单价（取自 T2.F_VBIK_JSDJ）
        // 说明：Excel G3 注释 —— "取小于等于结束日期的最近销售订单销售单价本位币字段"
        String basePriceStr = salesRow.get("销售单价");
        double basePriceVal = 0;
        try { basePriceVal = Double.parseDouble(basePriceStr); } catch (NumberFormatException e) {}
        row.put("基准价", basePriceVal);

        // 销售数量
        String qtyStr = salesRow.get("销售数量");
        double qty = 0;
        try { qty = Double.parseDouble(qtyStr); } catch (NumberFormatException e) {}
        row.put("销售数量", qty);

        row.put("客户", salesRow.get("客户"));

        reportRows.add(row);
    }

    } // end if (searched)
%>

    <%-- ==================== 主数据表格 ==================== --%>
    <table>
        <thead>
            <tr>
                <th>序号</th>
                <th style="min-width:140px;">型号</th>
                <th class="col-cu">铜用量<br/>(克)</th>
                <th class="col-al">铝用量<br/>(克)</th>
                <th class="col-sn">锡用量<br/>(克)</th>
                <th class="col-ni">镍用量<br/>(克)</th>
                <th class="col-cost">成本价汇总<br/>(元)</th>
                <th class="col-base">基准价<br/>(元)</th>
                <th class="col-qty">销售数量</th>
                <th>客户</th>
            </tr>
        </thead>
        <tbody>
<%
    int rowNum = 0;
    double grandCu = 0, grandAl = 0, grandSn = 0, grandNi = 0;
    double grandCost = 0, grandBase = 0, grandQty = 0;

    if (!searched) {
%>
            <tr>
                <td colspan="10" style="color:#999; padding: 32px; text-align: center;">
                    💡 请设置日期范围和客户筛选条件，然后点击 <b>"查询"</b> 按钮加载数据
                </td>
            </tr>
<%
    } else if (reportRows.isEmpty()) {
%>
            <tr>
                <td colspan="10" style="color:#999; padding: 24px;">
                    暂无数据（请调整筛选条件后重试）
                </td>
            </tr>
<%
    } else {
        for (Map<String, Object> row : reportRows) {
            rowNum++;
            double cu   = (Double) row.get("铜用量");
            double al   = (Double) row.get("铝用量");
            double sn   = (Double) row.get("锡用量");
            double ni   = (Double) row.get("镍用量");
            double cost = (Double) row.get("成本价汇总");
            double base = (Double) row.get("基准价");
            double qty  = (Double) row.get("销售数量");

            grandCu += cu;   grandAl += al;   grandSn += sn;   grandNi += ni;
            grandCost += cost;   grandBase += base;   grandQty += qty;

            boolean hasUsage = (cu > 0 || al > 0 || sn > 0 || ni > 0);
%>
            <tr>
                <td><%= rowNum %></td>
                <td class="left"><b><%= row.get("基础型号") %></b></td>
                <td class="num <%= cu > 0 ? "" : "zero" %>"><%= cu > 0 ? String.format("%.2f", cu) : "-" %></td>
                <td class="num <%= al > 0 ? "" : "zero" %>"><%= al > 0 ? String.format("%.2f", al) : "-" %></td>
                <td class="num <%= sn > 0 ? "" : "zero" %>"><%= sn > 0 ? String.format("%.2f", sn) : "-" %></td>
                <td class="num <%= ni > 0 ? "" : "zero" %>"><%= ni > 0 ? String.format("%.2f", ni) : "-" %></td>
                <td class="num"><%= hasUsage ? String.format("%.2f", cost) : "-" %></td>
                <td class="num"><%= hasUsage ? String.format("%.2f", base) : "-" %></td>
                <td class="num"><%= String.format("%.0f", qty) %></td>
                <td class="left" style="font-size:12px;"><%= row.get("客户") != null ? row.get("客户") : "" %></td>
            </tr>
<%
        }
    }
%>
        </tbody>
        <% if (!reportRows.isEmpty()) { %>
        <tfoot>
            <tr style="background-color: #fafafa; font-weight: bold;">
                <td colspan="2">合计</td>
                <td class="num"><%= String.format("%.2f", grandCu) %></td>
                <td class="num"><%= String.format("%.2f", grandAl) %></td>
                <td class="num"><%= String.format("%.2f", grandSn) %></td>
                <td class="num"><%= String.format("%.2f", grandNi) %></td>
                <td class="num"><%= String.format("%.2f", grandCost) %></td>
                <td class="num"><%= String.format("%.2f", grandBase) %></td>
                <td class="num"><%= String.format("%.0f", grandQty) %></td>
                <td></td>
            </tr>
        </tfoot>
        <% } %>
    </table>

    <%-- 金属现价参考说明栏 --%>
    <div class="price-bar">
        📊 <b>金属现价参考</b>（采购订单最新含税单价<% if (!filterDateTo.isEmpty()) { %> ≤ <%= filterDateTo %><% } %>）&nbsp;&nbsp;
        <%
            StringBuilder priceInfo = new StringBuilder();
            for (int j = 0; j < 4; j++) {
                double pricePerTon = currentPrice[j] * 1000000;
                priceInfo.append(metalNames[j]).append("：")
                         .append(String.format("%.0f", pricePerTon)).append(" 元/吨");
                if (j < 3) priceInfo.append("；");
            }
        %>
        <%= priceInfo.toString() %>
        &nbsp;&nbsp;|&nbsp;&nbsp;
        <b>基准单价</b>（元/吨）：铜 76,800；锡 256,900；铝 19,760；镍 133,500
        &nbsp;&nbsp;|&nbsp;&nbsp;
        <span style="font-size:12px;">成本价汇总 = Σ 用量(克) × 金属现价(元/克)；基准价 = 销售订单本位币单价(F_VBIK_JSDJ)</span>
    </div>

    <% if (!isExport) { %>

</body>
</html>
<% } %>
