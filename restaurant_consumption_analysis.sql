/*
PROJECT: Restaurant Outlet Consumption Analysis
SOURCE: DSR 24 DECEMBER - DAY 1.csv
DIALECT: Microsoft SQL Server (T-SQL)

IMPORTANT DATA NOTE
This source is a wide-format daily consumption sheet. It is not transaction-level
sales data. It supports quantity/consumption analysis, not revenue, profit,
discount, or sales conversion analysis unless those fields are joined from another source.

RECOMMENDED IMPORT
1. In Excel/Power Query, remove the top zone header row.
2. Use the second row as the header row; retain the item/category fields and outlet columns.
3. Fill category names down for item rows and remove subtotal/total rows.
4. Unpivot outlet columns so the final import has these columns:
   ReportDate, DayLabel, Category, ItemCode, ItemName, OutletName, Quantity
5. Import the unpivoted result into dbo.DSR_Consumption.
   For this file, set ReportDate to the actual date for "24 December" and DayLabel = 'DAY 1'.
   Do not infer the year from the filename; enter the correct year from your source.

This script creates the destination table and runs the analysis queries. Import the data
before running the analysis section.
*/

SET NOCOUNT ON;
GO

IF OBJECT_ID('dbo.DSR_Consumption', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.DSR_Consumption
    (
        ConsumptionID BIGINT IDENTITY(1,1) PRIMARY KEY,
        ReportDate    DATE           NOT NULL,
        DayLabel      NVARCHAR(30)   NULL,
        Category      NVARCHAR(150)  NULL,
        ItemCode      NVARCHAR(50)   NULL,
        ItemName      NVARCHAR(250)  NOT NULL,
        OutletName    NVARCHAR(150)  NOT NULL,
        Quantity      DECIMAL(18,3)  NULL,
        ImportedAt    DATETIME2      NOT NULL CONSTRAINT DF_DSR_Consumption_ImportedAt DEFAULT SYSUTCDATETIME()
    );
END;
GO

/* Optional: clear table before reloading the same dataset.
TRUNCATE TABLE dbo.DSR_Consumption;
GO
*/

/* ============================================================
   1. DATA QUALITY CHECKS
   ============================================================ */

-- Row count, date coverage, outlet coverage
SELECT
    COUNT_BIG(*) AS RowCount,
    MIN(ReportDate) AS FirstReportDate,
    MAX(ReportDate) AS LastReportDate,
    COUNT(DISTINCT OutletName) AS OutletCount,
    COUNT(DISTINCT ItemName) AS DistinctItems
FROM dbo.DSR_Consumption;

-- Missing or invalid values
SELECT *
FROM dbo.DSR_Consumption
WHERE ReportDate IS NULL
   OR NULLIF(LTRIM(RTRIM(ItemName)), '') IS NULL
   OR NULLIF(LTRIM(RTRIM(OutletName)), '') IS NULL
   OR Quantity IS NULL
   OR Quantity < 0
ORDER BY ReportDate, OutletName, ItemName;

-- Duplicate item/outlet/date records; investigate before aggregating
SELECT ReportDate, OutletName, ItemCode, ItemName, COUNT(*) AS DuplicateRows
FROM dbo.DSR_Consumption
GROUP BY ReportDate, OutletName, ItemCode, ItemName
HAVING COUNT(*) > 1
ORDER BY DuplicateRows DESC;

/* ============================================================
   2. OVERALL CONSUMPTION
   ============================================================ */

-- Most-consumed items across all outlets and dates
SELECT TOP (20)
    ItemName,
    SUM(Quantity) AS TotalQuantity,
    COUNT(DISTINCT OutletName) AS OutletsUsingItem,
    COUNT(DISTINCT ReportDate) AS DaysRecorded
FROM dbo.DSR_Consumption
GROUP BY ItemName
ORDER BY SUM(Quantity) DESC;

-- Least-consumed items with positive recorded quantity
SELECT TOP (20)
    ItemName,
    SUM(Quantity) AS TotalQuantity,
    COUNT(DISTINCT OutletName) AS OutletsUsingItem
FROM dbo.DSR_Consumption
GROUP BY ItemName
HAVING SUM(Quantity) > 0
ORDER BY SUM(Quantity) ASC;

-- Items with zero total consumption (only valid if zeroes are explicitly recorded)
SELECT ItemName, SUM(Quantity) AS TotalQuantity
FROM dbo.DSR_Consumption
GROUP BY ItemName
HAVING SUM(Quantity) = 0
ORDER BY ItemName;

/* ============================================================
   3. OUTLET / BRANCH COMPARISON
   ============================================================ */

-- Highest-consumption outlets overall
SELECT
    OutletName,
    SUM(Quantity) AS TotalQuantity,
    COUNT(DISTINCT ItemName) AS DistinctItems,
    COUNT(DISTINCT ReportDate) AS DaysRecorded,
    AVG(CAST(Quantity AS DECIMAL(18,3))) AS AvgQuantityPerRecord
FROM dbo.DSR_Consumption
GROUP BY OutletName
ORDER BY SUM(Quantity) DESC;

-- Lowest-consumption outlets overall
SELECT
    OutletName,
    SUM(Quantity) AS TotalQuantity,
    COUNT(DISTINCT ItemName) AS DistinctItems,
    COUNT(DISTINCT ReportDate) AS DaysRecorded
FROM dbo.DSR_Consumption
GROUP BY OutletName
ORDER BY SUM(Quantity) ASC;

-- Best and worst outlet for each item
WITH OutletItem AS
(
    SELECT OutletName, ItemName, SUM(Quantity) AS TotalQuantity
    FROM dbo.DSR_Consumption
    GROUP BY OutletName, ItemName
),
Ranked AS
(
    SELECT *,
        ROW_NUMBER() OVER (PARTITION BY ItemName ORDER BY TotalQuantity DESC, OutletName) AS BestRank,
        ROW_NUMBER() OVER (PARTITION BY ItemName ORDER BY TotalQuantity ASC, OutletName) AS WorstRank
    FROM OutletItem
)
SELECT ItemName, OutletName, TotalQuantity,
       CASE WHEN BestRank = 1 THEN 'Highest consumption'
            WHEN WorstRank = 1 THEN 'Lowest consumption' END AS Performance
FROM Ranked
WHERE BestRank = 1 OR WorstRank = 1
ORDER BY ItemName, Performance;

-- Outlet share of total recorded quantity (quantity share, NOT sales share)
WITH OutletTotals AS
(
    SELECT OutletName, SUM(Quantity) AS TotalQuantity
    FROM dbo.DSR_Consumption
    GROUP BY OutletName
)
SELECT OutletName, TotalQuantity,
       CAST(100.0 * TotalQuantity / NULLIF(SUM(TotalQuantity) OVER (), 0) AS DECIMAL(8,2)) AS QuantitySharePct
FROM OutletTotals
ORDER BY TotalQuantity DESC;

/* ============================================================
   4. CATEGORY ANALYSIS
   ============================================================ */

SELECT
    COALESCE(Category, 'Unclassified') AS Category,
    SUM(Quantity) AS TotalQuantity,
    COUNT(DISTINCT ItemName) AS DistinctItems,
    COUNT(DISTINCT OutletName) AS OutletsUsingCategory
FROM dbo.DSR_Consumption
GROUP BY COALESCE(Category, 'Unclassified')
ORDER BY SUM(Quantity) DESC;

-- Top items within each category
WITH CategoryItems AS
(
    SELECT COALESCE(Category, 'Unclassified') AS Category,
           ItemName, SUM(Quantity) AS TotalQuantity
    FROM dbo.DSR_Consumption
    GROUP BY COALESCE(Category, 'Unclassified'), ItemName
),
Ranked AS
(
    SELECT *, DENSE_RANK() OVER (PARTITION BY Category ORDER BY TotalQuantity DESC) AS ItemRank
    FROM CategoryItems
)
SELECT Category, ItemName, TotalQuantity, ItemRank
FROM Ranked
WHERE ItemRank <= 5
ORDER BY Category, ItemRank, ItemName;

/* ============================================================
   5. DAILY / WEEKLY / MONTHLY / YEARLY TRENDS
   ============================================================ */

-- Daily quantity by outlet
SELECT ReportDate, OutletName, SUM(Quantity) AS TotalQuantity
FROM dbo.DSR_Consumption
GROUP BY ReportDate, OutletName
ORDER BY ReportDate, TotalQuantity DESC;

-- Weekly quantity by outlet (week starts Monday)
SET DATEFIRST 1;
SELECT
    DATEADD(DAY, 1 - DATEPART(WEEKDAY, ReportDate), ReportDate) AS WeekStart,
    OutletName,
    SUM(Quantity) AS TotalQuantity
FROM dbo.DSR_Consumption
GROUP BY DATEADD(DAY, 1 - DATEPART(WEEKDAY, ReportDate), ReportDate), OutletName
ORDER BY WeekStart, TotalQuantity DESC;

-- Monthly trend
SELECT
    DATEFROMPARTS(YEAR(ReportDate), MONTH(ReportDate), 1) AS MonthStart,
    SUM(Quantity) AS TotalQuantity,
    COUNT(DISTINCT OutletName) AS ActiveOutlets,
    COUNT(DISTINCT ItemName) AS ItemsRecorded
FROM dbo.DSR_Consumption
GROUP BY DATEFROMPARTS(YEAR(ReportDate), MONTH(ReportDate), 1)
ORDER BY MonthStart;

-- Yearly trend
SELECT YEAR(ReportDate) AS ReportYear, SUM(Quantity) AS TotalQuantity,
       COUNT(DISTINCT OutletName) AS Outlets, COUNT(DISTINCT ItemName) AS Items
FROM dbo.DSR_Consumption
GROUP BY YEAR(ReportDate)
ORDER BY ReportYear;

-- Year-over-year comparison by month; meaningful when multiple years are loaded
WITH Monthly AS
(
    SELECT YEAR(ReportDate) AS ReportYear, MONTH(ReportDate) AS ReportMonth,
           SUM(Quantity) AS TotalQuantity
    FROM dbo.DSR_Consumption
    GROUP BY YEAR(ReportDate), MONTH(ReportDate)
),
Compared AS
(
    SELECT ReportYear, ReportMonth, TotalQuantity,
           LAG(TotalQuantity) OVER (PARTITION BY ReportMonth ORDER BY ReportYear) AS PreviousAvailableYearQuantity
    FROM Monthly
)
SELECT ReportYear, ReportMonth, TotalQuantity, PreviousAvailableYearQuantity,
       CAST(100.0 * (TotalQuantity - PreviousAvailableYearQuantity)
            / NULLIF(PreviousAvailableYearQuantity, 0) AS DECIMAL(10,2)) AS ChangePct
FROM Compared
ORDER BY ReportMonth, ReportYear;

/* ============================================================
   6. NEW / DISCONTINUED ITEM MONITORING
   ============================================================ */

-- First and last recorded date per item
SELECT ItemName, MIN(ReportDate) AS FirstSeenDate, MAX(ReportDate) AS LastSeenDate,
       SUM(Quantity) AS TotalQuantity, COUNT(DISTINCT ReportDate) AS DaysRecorded
FROM dbo.DSR_Consumption
GROUP BY ItemName
ORDER BY FirstSeenDate, ItemName;

-- Items first appearing on each date
WITH FirstSeen AS
(
    SELECT ItemName, MIN(ReportDate) AS FirstSeenDate
    FROM dbo.DSR_Consumption
    GROUP BY ItemName
)
SELECT f.FirstSeenDate, d.ItemName
FROM FirstSeen f
JOIN (SELECT DISTINCT ItemName FROM dbo.DSR_Consumption) d
  ON d.ItemName = f.ItemName
ORDER BY f.FirstSeenDate, d.ItemName;

/* ============================================================
   7. OUTLET × ITEM MATRIX
   ============================================================ */

SELECT
    OutletName,
    ItemName,
    SUM(Quantity) AS TotalQuantity
FROM dbo.DSR_Consumption
GROUP BY OutletName, ItemName
ORDER BY OutletName, TotalQuantity DESC;

/* ============================================================
   8. OPTIONAL PRICING / PROFIT ANALYSIS
   ============================================================

   Consumption quantity alone cannot calculate revenue or profit.
   If you obtain a reliable item cost/price file, create and populate:

   CREATE TABLE dbo.ItemPricing
   (
       ItemCode NVARCHAR(50) NULL,
       ItemName NVARCHAR(250) NOT NULL,
       UnitSellingPrice DECIMAL(18,2) NULL,
       UnitCost DECIMAL(18,2) NULL
   );

   The following query estimates consumption value and cost only if the
   unit definitions match and the prices/costs are valid for the report date.
   For actual restaurant sales and gross margin, use POS sales quantities,
   discounts, recipe costs, and time-effective pricing instead.
*/

IF OBJECT_ID('dbo.ItemPricing', 'U') IS NOT NULL
BEGIN
    SELECT
        d.ItemName,
        SUM(d.Quantity) AS TotalQuantity,
        SUM(d.Quantity * p.UnitSellingPrice) AS EstimatedValue,
        SUM(d.Quantity * p.UnitCost) AS EstimatedCost,
        SUM(d.Quantity * (p.UnitSellingPrice - p.UnitCost)) AS EstimatedGrossProfit,
        CAST(100.0 * SUM(d.Quantity * (p.UnitSellingPrice - p.UnitCost))
             / NULLIF(SUM(d.Quantity * p.UnitSellingPrice), 0) AS DECIMAL(8,2)) AS EstimatedMarginPct
    FROM dbo.DSR_Consumption d
    JOIN dbo.ItemPricing p
      ON (p.ItemCode IS NOT NULL AND p.ItemCode = d.ItemCode)
      OR (p.ItemCode IS NULL AND p.ItemName = d.ItemName)
    GROUP BY d.ItemName
    ORDER BY EstimatedMarginPct ASC;
END;
GO

/* ============================================================
   9. RECOMMENDED GITHUB PROJECT NOTES
   ============================================================

   Suggested repository title:
   Restaurant Daily Consumption & Outlet Performance Analysis (SQL Server)

   Suggested business questions:
   - Which items have the highest/lowest recorded consumption?
   - Which outlets record the highest/lowest total quantities?
   - Which categories contribute the most to recorded consumption?
   - How does consumption change by day, week, month, and year?
   - Which items are newly introduced or no longer recorded?
   - Where are the differences between outlets large enough to investigate?

   Interpretation caution:
   Quantities can have different units (pieces, kilograms, litres, packs).
   Do not sum unlike units to claim a meaningful operational score unless
   the source defines comparable units. For cross-outlet comparisons, prefer
   same-item comparisons or normalize quantities by covers/orders where available.
*/
