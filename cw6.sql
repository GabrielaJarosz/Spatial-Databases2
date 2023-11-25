DROP TABLE IF EXISTS AdventureWorksDW2019.dbo.stg_dimemp
SELECT  EMPLOYEEKEY, FIRSTNAME, LASTNAME, TITLE 
INTO AdventureWorksDW2019.dbo.stg_dimemp
FROM DimEmployee 
WHERE EMPLOYEEKEY >= 270 and EMPLOYEEKEY <= 275;



DROP TABLE IF EXISTS AdventureWorksDW2019.dbo.scd_dimemp;
CREATE TABLE AdventureWorksDW2019.dbo.scd_dimemp (
 EmployeeKey int ,
 FirstName nvarchar(50) not null,
 LastName nvarchar(50) not null,
 Title nvarchar(50),
 StartDate datetime,
 EndDate datetime,
);

INSERT INTO AdventureWorksDW2019.dbo.scd_dimemp (EmployeeKey, FirstName, LastName, Title) 
SELECT  EMPLOYEEKEY, FIRSTNAME, LASTNAME, TITLE
FROM AdventureWorksDW2019.dbo.stg_dimemp;