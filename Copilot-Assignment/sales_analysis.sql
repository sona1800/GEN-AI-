-- Copilot Assignment: Sales Data Analysis
-- Questions 1-3: SQL Analysis Queries

-- ============================================================================
-- QUESTION 1: Find total sales per city grouped by City
-- ============================================================================
-- Purpose: Returns the total purchase amount for each city, number of transactions, 
-- and average purchase amount per city

SELECT 
    City,
    SUM(PurchaseAmount) AS TotalSales,
    COUNT(*) AS NumberOfTransactions,
    AVG(PurchaseAmount) AS AveragePurchaseAmount
FROM SalesData
GROUP BY City
ORDER BY TotalSales DESC;

-- ============================================================================
-- QUESTION 2: Find the top 5 cities by revenue in descending order
-- ============================================================================
-- Purpose: Returns the top 5 cities with the highest total revenue, 
-- number of transactions, and average purchase amount

SELECT TOP 5
    City,
    SUM(PurchaseAmount) AS TotalRevenue,
    COUNT(*) AS NumberOfTransactions,
    AVG(PurchaseAmount) AS AveragePurchaseAmount
FROM SalesData
GROUP BY City
ORDER BY TotalRevenue DESC;

-- ============================================================================
-- QUESTION 3: Find customers with purchase amounts above the average purchase amount
-- ============================================================================
-- Purpose: Returns all customers whose purchase amount exceeds the average,
-- ordered by purchase amount in descending order

SELECT 
    CustomerID,
    Name,
    City,
    Age,
    PurchaseAmount,
    PurchaseDate,
    ROUND(PurchaseAmount - (SELECT AVG(PurchaseAmount) FROM SalesData), 2) AS AmountAboveAverage
FROM SalesData
WHERE PurchaseAmount > (SELECT AVG(PurchaseAmount) FROM SalesData)
ORDER BY PurchaseAmount DESC;