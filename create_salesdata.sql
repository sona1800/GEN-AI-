-- Create SalesData table
CREATE TABLE SalesData (
    CustomerID INT PRIMARY KEY,
    Name NVARCHAR(100) NOT NULL,
    Age INT NOT NULL,
    City NVARCHAR(50) NOT NULL,
    PurchaseAmount DECIMAL(10, 2) NOT NULL,
    PurchaseDate DATE NOT NULL
);

-- Insert 10,000 rows of random data
-- Note: This script uses SQL Server syntax. Adjust for your database system if needed.

DECLARE @Counter INT = 1;
DECLARE @FirstNames TABLE (Name NVARCHAR(50));
DECLARE @LastNames TABLE (Name NVARCHAR(50));
DECLARE @Cities TABLE (City NVARCHAR(50));

-- Populate sample first names
INSERT INTO @FirstNames VALUES 
('John'), ('Jane'), ('Michael'), ('Sarah'), ('David'), ('Emma'), ('Robert'), ('Olivia'),
('James'), ('Sophia'), ('William'), ('Ava'), ('Richard'), ('Isabella'), ('Joseph'), ('Mia'),
('Thomas'), ('Charlotte'), ('Christopher'), ('Amelia');

-- Populate sample last names
INSERT INTO @LastNames VALUES 
('Smith'), ('Johnson'), ('Williams'), ('Brown'), ('Jones'), ('Garcia'), ('Miller'), ('Davis'),
('Rodriguez'), ('Martinez'), ('Hernandez'), ('Lopez'), ('Gonzalez'), ('Wilson'), ('Anderson'),
('Thomas'), ('Taylor'), ('Moore'), ('Jackson'), ('Martin');

-- Populate sample cities
INSERT INTO @Cities VALUES 
('New York'), ('Los Angeles'), ('Chicago'), ('Houston'), ('Phoenix'), ('Philadelphia'),
('San Antonio'), ('San Diego'), ('Dallas'), ('San Jose'), ('Austin'), ('Jacksonville'),
('Fort Worth'), ('Columbus'), ('Indianapolis'), ('Charlotte'), ('San Francisco'), ('Seattle'),
('Denver'), ('Boston');

-- Insert 10,000 rows
WHILE @Counter <= 10000
BEGIN
    INSERT INTO SalesData (CustomerID, Name, Age, City, PurchaseAmount, PurchaseDate)
    VALUES (
        @Counter,
        (SELECT TOP 1 Name FROM @FirstNames ORDER BY NEWID()) + ' ' + 
        (SELECT TOP 1 Name FROM @LastNames ORDER BY NEWID()),
        CAST(RAND(CHECKSUM(NEWID())) * 60 + 18 AS INT),  -- Age between 18 and 78
        (SELECT TOP 1 City FROM @Cities ORDER BY NEWID()),
        CAST(RAND(CHECKSUM(NEWID())) * 9999.99 + 10 AS DECIMAL(10, 2)),  -- Purchase amount between 10 and 10,000
        DATEADD(DAY, CAST(RAND(CHECKSUM(NEWID())) * 365 AS INT), '2024-01-01')  -- Date in 2024
    );
    SET @Counter = @Counter + 1;
END;

-- Verify the data
SELECT COUNT(*) as TotalRecords FROM SalesData;
SELECT TOP 10 * FROM SalesData;

-- ============================================================================
-- ANALYTICAL QUERIES
-- ============================================================================

-- Query 1: Find total sales per city grouped by City
-- Returns the total purchase amount for each city
SELECT 
    City,
    SUM(PurchaseAmount) AS TotalSales,
    COUNT(*) AS NumberOfTransactions,
    AVG(PurchaseAmount) AS AveragePurchaseAmount
FROM SalesData
GROUP BY City
ORDER BY City;

-- Query 2: Find the top 5 cities by revenue in descending order
-- Returns the top 5 cities with highest total revenue
SELECT TOP 5
    City,
    SUM(PurchaseAmount) AS TotalRevenue,
    COUNT(*) AS NumberOfTransactions,
    AVG(PurchaseAmount) AS AveragePurchaseAmount
FROM SalesData
GROUP BY City
ORDER BY TotalRevenue DESC;

-- Query 3: Find customers with purchase amounts above the average purchase amount
-- Returns all customers whose purchase amount exceeds the average
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
