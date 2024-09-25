declare @SubRunID NVARCHAR(30)
set @SubRunID =    'R201906131219140006'
--set @SubRunID = @SubRunID_in
declare @ChecksRunIDVar NVARCHAR(30)
declare @DataRunIDVar NVARCHAR(30)
DECLARE @LoopCounter INT
DECLARE @MaxLoopCounter INT
DECLARE @vValue nvarchar(1000)
DECLARE @vExcelCellType INT
DECLARE @vFormula nvarchar(1000)
DECLARE @Result INT
DECLARE @vi INT
DECLARE @vQuery NVARCHAR(1000)
DECLARE @SQLFormula1 NVARCHAR(1000)
DECLARE @SQLFormula NVARCHAR(1000)
DECLARE @createtable  NVARCHAR(MAX);
DECLARE @columnlist NVARCHAR(MAX);
--set @ChecksRunIDVar =  '2018-01-15T17_46_20.1672130Z'
--set @DataRunIDVar =  '2018-01-15T17_46_20.1672130Z'

DECLARE @DynamicPivotQuery AS NVARCHAR(MAX)='';
DECLARE @celldata TABLE(SheetName varchar(255),ColumnID varchar(255),RowID varchar(255),FieldName varchar(255),CellValue varchar(max))
DECLARE @fieldnames TABLE(FieldName varchar(255))
declare @fieldnameslist AS NVARCHAR(MAX)
declare @ReferenceSheetName varchar(1024)
declare @ReferenceTextStart varchar(1024)


DECLARE @DataSpecRunID  AS NVARCHAR(MAX)='';
DECLARE @TemplateMapRunID  AS NVARCHAR(MAX)='';
DECLARE @DataMapRunID  AS NVARCHAR(MAX)='';
DECLARE @dag  AS NVARCHAR(MAX)='';

DECLARE @ReferenceRowID as Integer
DECLARE @ReferenceColumnID as Integer

-------
DECLARE @DataSubmissionRunID  AS NVARCHAR(MAX)='';
DECLARE @DataDrive AS NVARCHAR(MAX)='';
DECLARE @DataEnvironment  AS NVARCHAR(MAX)='';
DECLARE @DataConveyorStage  AS NVARCHAR(MAX)='';
DECLARE @DataSubject  AS NVARCHAR(MAX)='';
DECLARE @DataContract  AS NVARCHAR(MAX)='';
DECLARE @DataDate  AS NVARCHAR(MAX)='';
DECLARE @DataPMU  AS NVARCHAR(MAX)='';
DECLARE @DataFolder8  AS NVARCHAR(MAX)='';
DECLARE @DataFolder9  AS NVARCHAR(MAX)='';
DECLARE @DataFolder10  AS NVARCHAR(MAX)='';
----

DECLARE @TableDatabase AS NVARCHAR(MAX)='';
DECLARE @TableDatasetID AS NVARCHAR(MAX)='';
DECLARE @TableSchema AS NVARCHAR(MAX)='';
DECLARE @TableName AS NVARCHAR(MAX)='';


-----
declare @AnchorColumn int;
declare @AnchorRow int;

DECLARE @CurrentFormula nvarchar(1000);

DECLARE @CurrentKeyWord NVARCHAR(1000);

DECLARE @SQLCmnd NVARCHAR(max);
----


---

SELECT
@DataSubmissionRunID = [SubmissionRunID]
,@DataDrive = [Folder1]
,@DataEnvironment = [Folder2]
,@DataConveyorStage = [Folder3]
,@DataSubject= [Folder4]
,@DataContract = [Folder5]
,@DataDate = [Folder6]
,@DataPMU = [Folder7]
,@DataFolder8 = [Folder8]
,@DataFolder9 = [Folder9]
,@DataFolder10 = [Folder10]
,@dag  = [Folder7]
FROM [Process].[Files] where [SubmissionRunID] = @SubRunID; -- and Folder7 = @dag


--------------------------------
begin transaction;
drop table IF EXISTS Working.DataCellData;
drop table IF EXISTS Working.DataSpecCells;
drop table IF EXISTS Working.DataSetCells;
drop table IF EXISTS Working.SpecList; 
drop table IF EXISTS Working.DataSetDataSet;
drop table IF EXISTS Working.DataSetCells;
drop table IF EXISTS Working.DataSetColumns;
drop table IF EXISTS Working.DataMapCells;
drop table IF EXISTS Working.DataMapColumns;
DROP TABLE IF EXISTS Working.DataMapTable; 
DROP TABLE IF EXISTS Working.DataMapDataSet
commit transaction;

-------------------------------

begin transaction;

select SheetName,CellType,try_cast(ColumnID as Integer) as ColumnID,Try_cast([RowID] as Integer) as RowID,
cast(CellValue as varchar(max)) as CellValue into Working.DataCellData 
from [Data].[Cells]
where [SubmissionRunID] = @SubRunID;

select [SubmissionRunID],[Folder8],[FileName] into Working.SpecList
from
(select
[SubmissionRunID],
[Folder6],
[Folder8],
[FileName],
max(Folder8) over (Partition by [FileName]) LatestVersion
from [Process].[Files]
where Folder4 = '100_ReferenceData' and Folder6 = @DataPMU) g
where [Folder8] = LatestVersion;


select  @DataSpecRunID = [SubmissionRunID] from Working.SpecList where [FileName] = 'DataSpec.xlsx';
select  @DataMapRunID = [SubmissionRunID] from Working.SpecList where [FileName] = 'DataMap.xlsx';

select SheetName,try_cast(ColumnID as Integer) as ColumnID,Try_cast([RowID] as Integer) as RowID, cast(CellValue as varchar(max)) as CellValue
into Working.DataSpecCells 
from [Data].[Cells]
where [SubmissionRunID] = @DataSpecRunID;

------

select * into Working.DataSetCells from Working.DataSpecCells where SheetName = 'DataSet' and RowID > 1;
select * into Working.DataSetColumns from Working.DataSpecCells where SheetName = 'DataSet' and RowID = 1;
select cel.*, col.CellValue as FieldName into Working.DataSetDataSet from Working.DataSetCells cel, Working.DataSetColumns col where cel.ColumnID = col.ColumnID;

---------------------

DROP TABLE IF EXISTS Working.fieldnames; 

select FieldName into Working.fieldnames from Working.DataSetDataSet;

set  @fieldnameslist = NULL

SELECT @fieldnameslist = ISNULL(@fieldnameslist + ',','') + QUOTENAME(FieldName)
FROM (SELECT DISTINCT FieldName FROM Working.fieldnames) AS FieldName;

DROP TABLE if exists Working.celldata; 
DROP TABLE if exists Working.pivoteddata;

--select 'DataSetDataSet';

--select * from Working.DataSetDataSet;


select * into Working.celldata from Working.DataSetDataSet;

------- Do Pivot ----------

--Prepare the PIVOT query using the dynamic 
SET @DynamicPivotQuery = 
N'select * into Working.pivoteddata from (select RowID,FieldName,CellValue from Working.celldata) src
PIVOT(max(CellValue) for FieldName in ('+@fieldnameslist+')) piv;';
--Execute the Dynamic Pivot Query

--select @DynamicPivotQuery;

EXEC sp_executesql @DynamicPivotQuery;

---- Populate Pivoted Table

DROP TABLE IF EXISTS Working.DataSetTable; 

select * into Working.DataSetTable from Working.pivoteddata;

DROP TABLE IF EXISTS Working.pivoteddata;

-----------------------


select distinct @ReferenceSheetName = SheetName, @ReferenceTextStart = ReferenceTextStart from working.DataSetTable;

select @AnchorColumn = ColumnID, @AnchorRow = RowID from working.DataCellData
where
SheetName = @ReferenceSheetName and CellValue = @ReferenceTextStart;

commit transaction;

------  data map ---------
begin transaction;

select SheetName,try_cast(ColumnID as Integer) as ColumnID,Try_cast([RowID] as Integer) as RowID, cast(CellValue as varchar(max)) as CellValue
into Working.DataMapCells 
from [Data].[Cells]
where [SubmissionRunID] = @DataMapRunID and SheetName = 'DataMap' and RowID > 1;

select SheetName,try_cast(ColumnID as Integer) as ColumnID,Try_cast([RowID] as Integer) as RowID, cast(CellValue as varchar(max)) as CellValue
into Working.DataMapColumns 
from [Data].[Cells]
where [SubmissionRunID] = @DataMapRunID and SheetName = 'DataMap' and RowID = 1;

select cel.*, col.CellValue as FieldName into Working.DataMapDataSet from Working.DataMapCells cel, Working.DataMapColumns col where cel.ColumnID = col.ColumnID;

---------------------

DROP TABLE IF EXISTS Working.fieldnames; 

select FieldName into Working.fieldnames from Working.DataMapDataSet;

set  @fieldnameslist = NULL

SELECT @fieldnameslist = ISNULL(@fieldnameslist + ',','') + QUOTENAME(FieldName)
FROM (SELECT DISTINCT FieldName FROM Working.fieldnames) AS FieldName;

DROP TABLE if exists Working.celldata; 
DROP TABLE if exists Working.pivoteddata;

--select 'DataSetDataSet';

--select * from Working.DataSetDataSet;


select * into Working.celldata from Working.DataMapDataSet;

------- Do Pivot ----------

--Prepare the PIVOT query using the dynamic 
SET @DynamicPivotQuery = 
N'select * into Working.pivoteddata from (select RowID,FieldName,CellValue from Working.celldata) src
PIVOT(max(CellValue) for FieldName in ('+@fieldnameslist+')) piv;';
--Execute the Dynamic Pivot Query

--select @DynamicPivotQuery;

EXEC sp_executesql @DynamicPivotQuery;

---- Populate Pivoted Table

DROP TABLE IF EXISTS Working.DataMapTable; 

select * into Working.DataMapTable from Working.pivoteddata;

DROP TABLE IF EXISTS Working.pivoteddata;

--select FieldID, ColumnName, ColumnType from Working.DataMapTable

commit transaction;

-----------------------
------  data meta data ---------
begin transaction;
DROP TABLE if exists Working.MetaDataColumns;
DROP TABLE if exists Working.MetaDataCells;
DROP TABLE if exists Working.MetaDataDataSet;
DROP TABLE if exists Working.MetaDataTable;

select SheetName,try_cast(ColumnID as Integer) as ColumnID,Try_cast([RowID] as Integer) as RowID, cast(CellValue as varchar(max)) as CellValue
into Working.MetaDataCells 
from [Data].[Cells]
where [SubmissionRunID] = @DataMapRunID and SheetName = 'MetaData' and RowID > 1;

select SheetName,try_cast(ColumnID as Integer) as ColumnID,Try_cast([RowID] as Integer) as RowID, cast(CellValue as varchar(max)) as CellValue
into Working.MetaDataColumns 
from [Data].[Cells]
where [SubmissionRunID] = @DataMapRunID and SheetName = 'MetaData' and RowID = 1;

select cel.*, col.CellValue as FieldName into Working.MetaDataDataSet from Working.MetaDataCells cel, Working.MetaDataColumns col where cel.ColumnID = col.ColumnID;

---------------------

DROP TABLE IF EXISTS Working.fieldnames; 

select FieldName into Working.fieldnames from Working.MetaDataDataSet;

set  @fieldnameslist = NULL

SELECT @fieldnameslist = ISNULL(@fieldnameslist + ',','') + QUOTENAME(FieldName)
FROM (SELECT DISTINCT FieldName FROM Working.fieldnames) AS FieldName;

DROP TABLE if exists Working.celldata; 
DROP TABLE if exists Working.pivoteddata;

--select 'DataSetDataSet';

--select * from Working.DataSetDataSet;


select * into Working.celldata from Working.MetaDataDataSet;

------- Do Pivot ----------

--Prepare the PIVOT query using the dynamic 
SET @DynamicPivotQuery = 
N'select * into Working.pivoteddata from (select RowID,FieldName,CellValue from Working.celldata) src
PIVOT(max(CellValue) for FieldName in ('+@fieldnameslist+')) piv;';
--Execute the Dynamic Pivot Query

--select @DynamicPivotQuery;

EXEC sp_executesql @DynamicPivotQuery;

---- Populate Pivoted Table

DROP TABLE IF EXISTS Working.MetaDataTable; 

select * into Working.MetaDataTable from Working.pivoteddata;

DROP TABLE IF EXISTS Working.pivoteddata;

-------------------------------------------------
SELECT 
      @TableDatabase = [Database]
      ,@TableDatasetID = [DataSetID]
      ,@TableSchema = [Schema]
      ,@TableName = [TableName]
  FROM [Working].[MetaDataTable]

commit transaction;

-----------------------


------------------------

begin transaction;

DROP TABLE IF EXISTS Working.PublishDataSet; 

select
--ds.[RowID]
ds.[ColumnFromReferenceText]
,ds.[DataSetID]
,ds.[FieldDescription]
,ds.[FieldID] as FieldName
,ds.[FieldPosition]
,ds.[ReferenceTextEnd]
,ds.[ReferenceTextStart]
,ds.[RowFromReferenceText]
,ds.[SheetName]
,ds.[TableOrientation]
--,dc.[SheetName]
,dc.[CellType]
,dc.[ColumnID]
,dc.[RowID]
,dc.[CellValue]
into Working.PublishDataSet from
working.DataSetTable ds, working.DataCellData dc
where
dc.SheetName = ds.SheetName and
dc.ColumnID = ds.ColumnFromReferenceText + @AnchorColumn and dc.RowID > ds.RowFromReferenceText + @AnchorRow;


--------------------- prepare for pivot -------------------------------------------------
DROP TABLE IF EXISTS Working.fieldnames; 

select FieldName into Working.fieldnames from Working.PublishDataSet;

set  @fieldnameslist = NULL

SELECT @fieldnameslist = ISNULL(@fieldnameslist + ',','') + QUOTENAME(FieldName)
FROM (SELECT DISTINCT FieldName FROM Working.fieldnames) AS FieldName;

DROP TABLE if exists Working.celldata; 
DROP TABLE if exists Working.pivoteddata;

--select 'DataSetDataSet';

--select * from Working.DataSetDataSet;


select * into Working.celldata from Working.PublishDataSet;

------- Do Pivot ----------

--Prepare the PIVOT query using the dynamic 
SET @DynamicPivotQuery = 
N'select * into Working.pivoteddata from (select RowID,FieldName,CellValue from Working.celldata) src
PIVOT(max(CellValue) for FieldName in ('+@fieldnameslist+')) piv;';
--Execute the Dynamic Pivot Query

--select @DynamicPivotQuery;

EXEC sp_executesql @DynamicPivotQuery;

---- Populate Pivoted Table

DROP TABLE IF EXISTS Working.PublishTable; 


--------
set @createtable = NULL;

SELECT @createtable = ISNULL(@createtable + ',','') + CreateField
FROM (SELECT '['+ColumnName+']' + ' ' + ColumnType as CreateField FROM [Working].[DataMapTable]) AS CreateField;

SELECT @columnlist = ISNULL(@columnlist + ',','') + FieldName
FROM (SELECT '['+ColumnName+']' FieldName FROM [Working].[DataMapTable]) AS FieldName;

--SELECT ColumnName + ' ' + ColumnType as CreateField FROM [Working].[DataMapTable]

-- @TableSchema = [Schema]
-- @TableName = [TableName]

set @createtable = 'create table ' + @TableSchema +'.' + @TableName + '(' +@createtable + ')'

--select @createtable;

set @SQLCmnd = 'DROP TABLE IF EXISTS ' + @TableSchema +'.' + @TableName

EXEC sp_executesql @SQLCmnd;

EXEC sp_executesql @createtable;

set @SQLCmnd = 'insert into ' + @TableSchema +'.' + @TableName + ' select ' + @columnlist + ' from Working.pivoteddata;'

--select @SQLCmnd

EXEC sp_executesql @SQLCmnd;

DROP TABLE IF EXISTS Working.pivoteddata;

commit transaction;

--select * from Working.PublishTable;