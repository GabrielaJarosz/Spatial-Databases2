--b
WITH ranked AS (
    SELECT
        ProductKey,
        UnitPrice,
		OrderDate,
		OrderQuantity,
        ROW_NUMBER() OVER (PARTITION BY OrderDate ORDER BY UnitPrice DESC) AS RowNumber
    FROM
        dbo.FactInternetSales
)
SELECT
	ProductKey,
    OrderDate,
    UnitPrice
FROM
    ranked
WHERE
    RowNumber <= 3
ORDER BY OrderDate;