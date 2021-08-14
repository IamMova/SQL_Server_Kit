/***************************************************************************************************
Script Name			:   Excel_To_SQLServer_Table
Create Date			:	2016-08-29
Author				:   Mova
Description			:   This script is helpful when DB Developer wants to dump the data from Excel to
						any SQL Server table without using import/export feature.
						This script will just prepare Excel formula for you and you just need to paste
						it in Excel cell, which will create a SQL Query.
Call by				:	NA
Affected table(s)	:	NA
Used By				:	Used by Database Developer, where they want to dump the data from Excel to 
						SQL Server Table
Parameter(s)/Values	:   @TableName			- Name of the Table to be created
						@CreateActualTable	- Can be Yes/No. Use Yes, only if want to create and insert the data in actual table
						@SheetNameOfHeader	- Excel sheet name for header (table column list)
						@HeaderColumnRow	- To Specify the Row # from where the header can be pulled
						@DataColumnRow		- To Specify the Row # from where the data starts
						@StartingColumn		- Start range of excel column
						@EndingColumn		- End range of excel column
						@SeparatedDelimiter - This will be only changed you are using delimiter other than comma(,)
Usage				:	@TableName			= 'Order_Details_2020'
						@CreateActualTable	= 'No'
						@SheetNameOfHeader	= 'Orders Of 2020' 
						@HeaderColumnRow	= '1' 
						@DataColumnRow		= '2' 
						@StartingColumn		= 'A'
						@EndingColumn		= 'G'
Source				:	https://github.com/IamMova/SQLServerPackage/ExcelPack
Sample File			:	Sample_Data_File.xlsx
****************************************************************************************************
SUMMARY OF CHANGES
Date(yyyy-mm-dd)    Version/Ticket #	Author              Comments
-------------------------------------------------------------------------------------------------------
2016-08-29			1.0					Mova				Created
***************************************************************************************************/
-- Setup (Start)
DECLARE @TableName NVARCHAR(MAX)		= 'Order_Details_2020'	-- Name of Table you want to create or use
DECLARE @CreateActualTable VARCHAR(10)	= 'No'					-- Yes | No [*Yes - Refer Actual Table *No - Create Temp Table]
DECLARE @SheetNameOfHeader VARCHAR(100) = 'Orders Of 2020'		-- Name of Excel sheet to get the column name 
DECLARE @HeaderColumnRow VARCHAR(100)	= '1'					-- Specify Row # of Header Column
DECLARE @DataColumnRow VARCHAR(100)		= '2'					-- Specify the Row # from where the data starts
DECLARE @StartingColumn NVARCHAR(100)	= 'A'					-- Specify the start column
DECLARE @EndingColumn NVARCHAR(100)		= 'G'					-- Specify the end column
DECLARE @SeparatedDelimiter VARCHAR(10) = ','					-- In many cases this will be comma. but if you are using different seperator, specify it here
-- Setup (End)

/***************************************************************************************************/
-- **Don't touch anything from here until and unless you know what you are doing

DECLARE @StartingColumnINT INT = ASCII(@StartingColumn)
DECLARE @EndingColumnINT INT = ASCII(@EndingColumn)

DECLARE @ListOfColumns AS TABLE (COLUMNNAME NVARCHAR(100))

DECLARE @SelectFormula NVARCHAR(MAX) = ''
DECLARE @CreateTableFormula NVARCHAR(MAX) = ''
DECLARE @InsertFormula NVARCHAR(MAX) = ''
DECLARE @TableExitstFormula NVARCHAR(MAX) = ''
DECLARE @ExcelColumn VARCHAR(MAX)= ''

;SET NOCOUNT ON;

IF OBJECT_ID('tempdb..#Number_Combination') IS NOT NULL
BEGIN
	DROP TABLE #Number_Combination
END

;WITH NUMBERS (N)
AS (SELECT N FROM(VALUES('A'),('B'),('C'),('D'),('E'),('F'),('G'),('H'),('I'),('J'),('K'),('L'),('M'),('N'),('O'),('P'),('Q'),('R'),('S'),('T'),('U'),('V'),('W'),('X'),('Y'),('Z')) NUMBERS(N))
, RECUR 
(
	N, COMBINATION
)
AS (
	SELECT N
		, CAST(N AS VARCHAR(1000))
	FROM NUMBERS
	
	UNION ALL
	
	SELECT N.N
		, CAST(R.COMBINATION + CAST(N.N AS VARCHAR(10)) AS VARCHAR(1000))
	FROM RECUR R
	INNER JOIN NUMBERS N ON N.N >= R.N OR N.N <= R.N
	WHERE LEN(CAST(R.COMBINATION + CAST(N.N AS VARCHAR(10)) AS VARCHAR(1000))) <= 3
	)
SELECT ROW_NUMBER() OVER(ORDER BY LEN(COMBINATION), COMBINATION) N_ORDER, COMBINATION
INTO #Number_Combination
FROM RECUR
ORDER BY LEN(COMBINATION), COMBINATION;

SELECT @StartingColumnINT = N_ORDER FROM #Number_Combination WHERE COMBINATION = @StartingColumn
SELECT @EndingColumnINT = N_ORDER FROM #Number_Combination WHERE COMBINATION = @EndingColumn

IF ISNULL(@EndingColumnINT, 0) > ISNULL(@StartingColumnINT, 0) AND ISNULL(@StartingColumnINT, 0) != 0 AND ISNULL(@EndingColumnINT, 0) != 0
BEGIN
	WHILE(@StartingColumnINT<=@EndingColumnINT)
	BEGIN
		   SELECT @ExcelColumn = COMBINATION FROM #Number_Combination WHERE N_ORDER = @StartingColumnINT

		   SELECT @SelectFormula = @SelectFormula + ', ''"'+@SeparatedDelimiter+'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE('+@ExcelColumn+''+@DataColumnRow+', "''", "''''"), CHAR(13), "''+CHAR(13)+''"), CHAR(10),"''+CHAR(10)+''"),"","''"),"'''
		   SELECT @CreateTableFormula = @CreateTableFormula +', ["'+@SeparatedDelimiter+''''+@SheetNameOfHeader+'''!$'+@ExcelColumn+'$'+@HeaderColumnRow+''+@SeparatedDelimiter+'"] NVARCHAR(MAX)'
		   SELECT @InsertFormula = @InsertFormula + ', ["'+@SeparatedDelimiter+''''+@SheetNameOfHeader+'''!$'+@ExcelColumn+'$'+@HeaderColumnRow+''+@SeparatedDelimiter+'"]'
       
		   SET @StartingColumnINT = @StartingColumnINT+1  
		   SET @ExcelColumn = '' 
	END

	SELECT @SelectFormula = @SeparatedDelimiter+'" VALUES('+STUFF(@SelectFormula, 1,2, '')+')"'

	IF ISNULL(@CreateActualTable, 'NO') = 'NO'
	BEGIN
		   SELECT @InsertFormula = '=CONCATENATE("INSERT INTO #'+@TableName+'('+STUFF(@InsertFormula, 1, 2, '')+')"'+@SelectFormula+')'
		   SELECT @CreateTableFormula = '=CONCATENATE("CREATE TABLE #'+@TableName+' (ID INT IDENTITY(1,1), '+STUFF(@CreateTableFormula, 1,2,'')+')")'	   
		   SELECT @TableExitstFormula = 'IF OBJECT_ID(''tempdb..#'+@TableName+''') IS NOT NULL'+CHAR(10)+'BEGIN'+CHAR(10)+CHAR(9)+'DROP TABLE #'+@TableName+CHAR(10)+'END'
		   PRINT @TableExitstFormula	   
	END
	ELSE 
	BEGIN
		   SELECT @InsertFormula = '=CONCATENATE("INSERT INTO '+@TableName+'('+STUFF(@InsertFormula, 1, 2, '')+')"'+@SelectFormula+')'
		   SELECT @CreateTableFormula = '=CONCATENATE("CREATE TABLE '+@TableName+' (ID INT IDENTITY(1,1), '+STUFF(@CreateTableFormula, 1,2,'')+')")'
	END

	SELECT @CreateTableFormula AS [Create Table Query], @InsertFormula AS [Insert Into Query]
END
ELSE
BEGIN
	PRINT 'Error! Please select proper starting and ending column name!'
END

;SET NOCOUNT OFF;