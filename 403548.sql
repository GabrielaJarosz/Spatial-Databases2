CREATE PROCEDURE GetCurrencyRates
    @YearsAgo INT
AS
BEGIN
    DECLARE @CutoffDate DATE;
    SET @CutoffDate = DATEADD(YEAR, -@YearsAgo, GETDATE());
    SELECT cr.*
    FROM FactCurrencyRate cr
    INNER JOIN DimCurrency dc ON cr.CurrencyKey = dc.CurrencyKey
    WHERE (dc.CurrencyAlternateKey = 'GBP' OR dc.CurrencyAlternateKey = 'EUR')
    AND cr.Date <= @CutoffDate;
END
