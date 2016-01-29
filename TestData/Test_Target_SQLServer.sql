-- Prepare to remove foreign keys
DECLARE @FkName  SYSNAME

-- Drop foreign keys for table 'AllDataTypes'
DECLARE fk_cursor CURSOR FOR 
  SELECT CONSTRAINT_NAME
    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
   WHERE TABLE_SCHEMA = 'dbo'
     AND TABLE_NAME = 'AllDataTypes'
     AND CONSTRAINT_TYPE = 'FOREIGN KEY'
ORDER BY CONSTRAINT_NAME

OPEN fk_cursor
FETCH NEXT FROM fk_cursor INTO @FkName
WHILE @@FETCH_STATUS = 0
BEGIN
  EXEC('ALTER TABLE [AllDataTypes] DROP CONSTRAINT ' + @FkName)
  FETCH NEXT FROM fk_cursor INTO @FkName
END

CLOSE fk_cursor
DEALLOCATE fk_cursor
;

-- Drop foreign keys for table 'Departments'
DECLARE fk_cursor CURSOR FOR 
  SELECT CONSTRAINT_NAME
    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
   WHERE TABLE_SCHEMA = 'dbo'
     AND TABLE_NAME = 'Departments'
     AND CONSTRAINT_TYPE = 'FOREIGN KEY'
ORDER BY CONSTRAINT_NAME

OPEN fk_cursor
FETCH NEXT FROM fk_cursor INTO @FkName
WHILE @@FETCH_STATUS = 0
BEGIN
  EXEC('ALTER TABLE [Departments] DROP CONSTRAINT ' + @FkName)
  FETCH NEXT FROM fk_cursor INTO @FkName
END

CLOSE fk_cursor
DEALLOCATE fk_cursor
;

-- Drop foreign keys for table 'Employees'
DECLARE fk_cursor CURSOR FOR 
  SELECT CONSTRAINT_NAME
    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
   WHERE TABLE_SCHEMA = 'dbo'
     AND TABLE_NAME = 'Employees'
     AND CONSTRAINT_TYPE = 'FOREIGN KEY'
ORDER BY CONSTRAINT_NAME

OPEN fk_cursor
FETCH NEXT FROM fk_cursor INTO @FkName
WHILE @@FETCH_STATUS = 0
BEGIN
  EXEC('ALTER TABLE [Employees] DROP CONSTRAINT ' + @FkName)
  FETCH NEXT FROM fk_cursor INTO @FkName
END

CLOSE fk_cursor
DEALLOCATE fk_cursor
;

-- Drop foreign keys for table 'ItemBranches'
DECLARE fk_cursor CURSOR FOR 
  SELECT CONSTRAINT_NAME
    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
   WHERE TABLE_SCHEMA = 'dbo'
     AND TABLE_NAME = 'ItemBranches'
     AND CONSTRAINT_TYPE = 'FOREIGN KEY'
ORDER BY CONSTRAINT_NAME

OPEN fk_cursor
FETCH NEXT FROM fk_cursor INTO @FkName
WHILE @@FETCH_STATUS = 0
BEGIN
  EXEC('ALTER TABLE [ItemBranches] DROP CONSTRAINT ' + @FkName)
  FETCH NEXT FROM fk_cursor INTO @FkName
END

CLOSE fk_cursor
DEALLOCATE fk_cursor
;

-- Drop foreign keys for table 'Items'
DECLARE fk_cursor CURSOR FOR 
  SELECT CONSTRAINT_NAME
    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
   WHERE TABLE_SCHEMA = 'dbo'
     AND TABLE_NAME = 'Items'
     AND CONSTRAINT_TYPE = 'FOREIGN KEY'
ORDER BY CONSTRAINT_NAME

OPEN fk_cursor
FETCH NEXT FROM fk_cursor INTO @FkName
WHILE @@FETCH_STATUS = 0
BEGIN
  EXEC('ALTER TABLE [Items] DROP CONSTRAINT ' + @FkName)
  FETCH NEXT FROM fk_cursor INTO @FkName
END

CLOSE fk_cursor
DEALLOCATE fk_cursor
;

-- Drop foreign keys for table 'TestForeignKeyOptions'
DECLARE fk_cursor CURSOR FOR 
  SELECT CONSTRAINT_NAME
    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
   WHERE TABLE_SCHEMA = 'dbo'
     AND TABLE_NAME = 'TestForeignKeyOptions'
     AND CONSTRAINT_TYPE = 'FOREIGN KEY'
ORDER BY CONSTRAINT_NAME

OPEN fk_cursor
FETCH NEXT FROM fk_cursor INTO @FkName
WHILE @@FETCH_STATUS = 0
BEGIN
  EXEC('ALTER TABLE [TestForeignKeyOptions] DROP CONSTRAINT ' + @FkName)
  FETCH NEXT FROM fk_cursor INTO @FkName
END

CLOSE fk_cursor
DEALLOCATE fk_cursor
;

-- Drop foreign keys for table 'TestForeignKeyOptions2'
DECLARE fk_cursor CURSOR FOR 
  SELECT CONSTRAINT_NAME
    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
   WHERE TABLE_SCHEMA = 'dbo'
     AND TABLE_NAME = 'TestForeignKeyOptions2'
     AND CONSTRAINT_TYPE = 'FOREIGN KEY'
ORDER BY CONSTRAINT_NAME

OPEN fk_cursor
FETCH NEXT FROM fk_cursor INTO @FkName
WHILE @@FETCH_STATUS = 0
BEGIN
  EXEC('ALTER TABLE [TestForeignKeyOptions2] DROP CONSTRAINT ' + @FkName)
  FETCH NEXT FROM fk_cursor INTO @FkName
END

CLOSE fk_cursor
DEALLOCATE fk_cursor
;

-- Drop foreign keys for table 'TestForeignKeyOptions3'
DECLARE fk_cursor CURSOR FOR 
  SELECT CONSTRAINT_NAME
    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
   WHERE TABLE_SCHEMA = 'dbo'
     AND TABLE_NAME = 'TestForeignKeyOptions3'
     AND CONSTRAINT_TYPE = 'FOREIGN KEY'
ORDER BY CONSTRAINT_NAME

OPEN fk_cursor
FETCH NEXT FROM fk_cursor INTO @FkName
WHILE @@FETCH_STATUS = 0
BEGIN
  EXEC('ALTER TABLE [TestForeignKeyOptions3] DROP CONSTRAINT ' + @FkName)
  FETCH NEXT FROM fk_cursor INTO @FkName
END

CLOSE fk_cursor
DEALLOCATE fk_cursor
;

-- Drop foreign keys for table 'TestForeignKeyOptions4'
DECLARE fk_cursor CURSOR FOR 
  SELECT CONSTRAINT_NAME
    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
   WHERE TABLE_SCHEMA = 'dbo'
     AND TABLE_NAME = 'TestForeignKeyOptions4'
     AND CONSTRAINT_TYPE = 'FOREIGN KEY'
ORDER BY CONSTRAINT_NAME

OPEN fk_cursor
FETCH NEXT FROM fk_cursor INTO @FkName
WHILE @@FETCH_STATUS = 0
BEGIN
  EXEC('ALTER TABLE [TestForeignKeyOptions4] DROP CONSTRAINT ' + @FkName)
  FETCH NEXT FROM fk_cursor INTO @FkName
END

CLOSE fk_cursor
DEALLOCATE fk_cursor
;

-- Drop foreign keys for table 'ZipCodes'
DECLARE fk_cursor CURSOR FOR 
  SELECT CONSTRAINT_NAME
    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS
   WHERE TABLE_SCHEMA = 'dbo'
     AND TABLE_NAME = 'ZipCodes'
     AND CONSTRAINT_TYPE = 'FOREIGN KEY'
ORDER BY CONSTRAINT_NAME

OPEN fk_cursor
FETCH NEXT FROM fk_cursor INTO @FkName
WHILE @@FETCH_STATUS = 0
BEGIN
  EXEC('ALTER TABLE [ZipCodes] DROP CONSTRAINT ' + @FkName)
  FETCH NEXT FROM fk_cursor INTO @FkName
END

CLOSE fk_cursor
DEALLOCATE fk_cursor
;

-- Drop table 'AllDataTypes'
IF EXISTS (
  SELECT * FROM INFORMATION_SCHEMA.TABLES
  WHERE TABLE_TYPE = 'BASE TABLE'
  AND TABLE_SCHEMA = 'dbo'
  AND TABLE_NAME = 'AllDataTypes'
  )
BEGIN
  DROP TABLE [dbo].[AllDataTypes]
END
;

-- Drop table 'Departments'
IF EXISTS (
  SELECT * FROM INFORMATION_SCHEMA.TABLES
  WHERE TABLE_TYPE = 'BASE TABLE'
  AND TABLE_SCHEMA = 'dbo'
  AND TABLE_NAME = 'Departments'
  )
BEGIN
  DROP TABLE [dbo].[Departments]
END
;

-- Drop table 'Employees'
IF EXISTS (
  SELECT * FROM INFORMATION_SCHEMA.TABLES
  WHERE TABLE_TYPE = 'BASE TABLE'
  AND TABLE_SCHEMA = 'dbo'
  AND TABLE_NAME = 'Employees'
  )
BEGIN
  DROP TABLE [dbo].[Employees]
END
;

-- Drop table 'ItemBranches'
IF EXISTS (
  SELECT * FROM INFORMATION_SCHEMA.TABLES
  WHERE TABLE_TYPE = 'BASE TABLE'
  AND TABLE_SCHEMA = 'dbo'
  AND TABLE_NAME = 'ItemBranches'
  )
BEGIN
  DROP TABLE [dbo].[ItemBranches]
END
;

-- Drop table 'Items'
IF EXISTS (
  SELECT * FROM INFORMATION_SCHEMA.TABLES
  WHERE TABLE_TYPE = 'BASE TABLE'
  AND TABLE_SCHEMA = 'dbo'
  AND TABLE_NAME = 'Items'
  )
BEGIN
  DROP TABLE [dbo].[Items]
END
;

-- Drop table 'TestForeignKeyOptions'
IF EXISTS (
  SELECT * FROM INFORMATION_SCHEMA.TABLES
  WHERE TABLE_TYPE = 'BASE TABLE'
  AND TABLE_SCHEMA = 'dbo'
  AND TABLE_NAME = 'TestForeignKeyOptions'
  )
BEGIN
  DROP TABLE [dbo].[TestForeignKeyOptions]
END
;

-- Drop table 'TestForeignKeyOptions2'
IF EXISTS (
  SELECT * FROM INFORMATION_SCHEMA.TABLES
  WHERE TABLE_TYPE = 'BASE TABLE'
  AND TABLE_SCHEMA = 'dbo'
  AND TABLE_NAME = 'TestForeignKeyOptions2'
  )
BEGIN
  DROP TABLE [dbo].[TestForeignKeyOptions2]
END
;

-- Drop table 'TestForeignKeyOptions3'
IF EXISTS (
  SELECT * FROM INFORMATION_SCHEMA.TABLES
  WHERE TABLE_TYPE = 'BASE TABLE'
  AND TABLE_SCHEMA = 'dbo'
  AND TABLE_NAME = 'TestForeignKeyOptions3'
  )
BEGIN
  DROP TABLE [dbo].[TestForeignKeyOptions3]
END
;

-- Drop table 'TestForeignKeyOptions4'
IF EXISTS (
  SELECT * FROM INFORMATION_SCHEMA.TABLES
  WHERE TABLE_TYPE = 'BASE TABLE'
  AND TABLE_SCHEMA = 'dbo'
  AND TABLE_NAME = 'TestForeignKeyOptions4'
  )
BEGIN
  DROP TABLE [dbo].[TestForeignKeyOptions4]
END
;

-- Drop table 'ZipCodes'
IF EXISTS (
  SELECT * FROM INFORMATION_SCHEMA.TABLES
  WHERE TABLE_TYPE = 'BASE TABLE'
  AND TABLE_SCHEMA = 'dbo'
  AND TABLE_NAME = 'ZipCodes'
  )
BEGIN
  DROP TABLE [dbo].[ZipCodes]
END
;

--  Create table 'AllDataTypes'
CREATE TABLE [AllDataTypes] (
   [DataTypeID] int IDENTITY (1,1) NOT NULL
  ,[DataTypeName] nvarchar(15) NOT NULL
  ,[DtBigint] bigint NOT NULL CONSTRAINT DF_AllDataTypes_DtBigint DEFAULT (1)
  ,[DtNumeric] numeric NOT NULL CONSTRAINT DF_AllDataTypes_DtNumeric DEFAULT (2.2)
  ,[DtNumeric_8_2] numeric(8, 2) NOT NULL CONSTRAINT DF_AllDataTypes_DtNumeric_8_2 DEFAULT (3.3)
  ,[Dtbit] bit NOT NULL CONSTRAINT DF_AllDataTypes_Dtbit DEFAULT (0)
  ,[DtSmallint] smallint NOT NULL CONSTRAINT DF_AllDataTypes_DtSmallint DEFAULT (5)
  ,[DtDecimal] decimal NOT NULL CONSTRAINT DF_AllDataTypes_DtDecimal DEFAULT (6.6)
  ,[DtDecimal_10_2] decimal(10, 2) NOT NULL CONSTRAINT DF_AllDataTypes_DtDecimal_10_2 DEFAULT (7.7)
  ,[DtSmallMoney] smallmoney NOT NULL CONSTRAINT DF_AllDataTypes_DtSmallMoney DEFAULT (8.8)
  ,[DtInt] int NOT NULL CONSTRAINT DF_AllDataTypes_DtInt DEFAULT (9)
  ,[DtTinyInt] tinyint NOT NULL
  ,[DtMoney] money NOT NULL
  ,[DtFloat] float NOT NULL
  ,[DtReal] real NOT NULL
  ,[DtDate] date NOT NULL
  ,[DtDatetimeOffset] datetimeoffset NOT NULL
  ,[DtDatetime2] datetime2 NOT NULL
  ,[DtSmallDatetime] smalldatetime NOT NULL
  ,[DtDatetime] datetime NOT NULL
  ,[DtTime] time NOT NULL
  ,[DtChar] char(255) NULL
  ,[DtVarchar] varchar(255) NOT NULL CONSTRAINT DF_AllDataTypes_DtVarchar DEFAULT ('')
  ,[DtVarcharMax] varchar(max) NOT NULL CONSTRAINT DF_AllDataTypes_DtVarcharMax DEFAULT ('')
  ,[DtNchar] nchar(255) NULL
  ,[DtNvarchar] nvarchar(255) NOT NULL CONSTRAINT DF_AllDataTypes_DtNvarchar DEFAULT ('A')
  ,[DtNvarcharMax] nvarchar(max) NOT NULL CONSTRAINT DF_AllDataTypes_DtNvarcharMax DEFAULT ('B')
  ,[DtBinary] binary(8) NULL
  ,[DtVarbinary] varbinary(1000) NULL
  ,[DtVarbinaryMax] varbinary(max) NULL
  ,[DtTimestamp] timestamp NOT NULL
  ,[DtHierarchyid] hierarchyid NOT NULL
  ,[DtUniqueIdentifier] uniqueidentifier NOT NULL CONSTRAINT DF_AllDataTypes_DtUniqueIdentifier DEFAULT (newid())
  ,[DtSqlVariant] sql_variant NULL
  ,[DtXml] xml NULL
  ,[DtGeography] geography NULL
  ,[DtGeometry] geometry NULL
  ,CONSTRAINT PK_AllDataTypes PRIMARY KEY NONCLUSTERED (DataTypeID)
  ,CONSTRAINT IK_AllDataTypes_DataTypeName UNIQUE CLUSTERED (DataTypeName)
)

EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Sample table for most common data types'' definitions.'
  , N'user', N'dbo', N'table', N'AllDataTypes'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'(Label: Data Type ID)'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DataTypeID'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Test single quatation (Label: Data Type''s Name)'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DataTypeName'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Exact'' Numerics'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtBigint'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Exact Numerics'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtNumeric'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Exact Numerics'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtNumeric_8_2'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Exact Numerics'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'Dtbit'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Exact Numerics'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtSmallint'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Exact Numerics'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtDecimal'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Exact Numerics'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtDecimal_10_2'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Exact Numerics'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtSmallMoney'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Exact Numerics'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtInt'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Exact Numerics'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtTinyInt'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Exact Numerics'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtMoney'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Approximate Numerics'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtFloat'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Approximate Numerics'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtReal'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Date and Time'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtDate'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Date and Time'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtDatetimeOffset'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Date and Time'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtDatetime2'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Date and Time'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtSmallDatetime'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Date and Time'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtDatetime'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Date and Time'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtTime'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Character Strings'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtChar'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Character Strings'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtVarchar'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Character Strings'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtVarcharMax'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Unicode Character Strings'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtNchar'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Unicode Character Strings'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtNvarchar'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Unicode Character Strings'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtNvarcharMax'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Binary Strings'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtBinary'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Binary Strings'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtVarbinary'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Binary Strings'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtVarbinaryMax'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Other Data Types'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtTimestamp'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Other Data Types'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtHierarchyid'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Other Data Types'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtUniqueIdentifier'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Other Data Types'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtSqlVariant'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Other Data Types'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtXml'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Spatial Types'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtGeography'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Spatial Types'
  , N'user', N'dbo', N'table', N'AllDataTypes', N'column', N'DtGeometry'
;

--  Create table 'Departments'
CREATE TABLE [Departments] (
   [DepartmentID] int IDENTITY (1,1) NOT NULL
  ,[DepartmentName] nvarchar(50) NOT NULL
  ,[ParentID] int NULL
  ,[ManagerID] int NULL
  ,CONSTRAINT PK_Departments PRIMARY KEY CLUSTERED (DepartmentID)
)

EXECUTE sp_addextendedproperty N'MS_Description'
  , N'The department table.'
  , N'user', N'dbo', N'table', N'Departments'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'(Label: Department ID)'
  , N'user', N'dbo', N'table', N'Departments', N'column', N'DepartmentID'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'(Label: Department Name)'
  , N'user', N'dbo', N'table', N'Departments', N'column', N'DepartmentName'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'(Label: Parent Department)'
  , N'user', N'dbo', N'table', N'Departments', N'column', N'ParentID'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'(Label: Manager)'
  , N'user', N'dbo', N'table', N'Departments', N'column', N'ManagerID'
;

--  Create table 'Employees'
CREATE TABLE [Employees] (
   [EmployeeID] int IDENTITY (1,1) NOT NULL
  ,[LastName] nvarchar(50) NOT NULL
  ,[FirstName] nvarchar(50) NOT NULL
  ,[DepartmentID] int NOT NULL
  ,CONSTRAINT PK_Employees PRIMARY KEY NONCLUSTERED (EmployeeID)
)
CREATE CLUSTERED INDEX IK_Employees_FirstName_LastName ON [Employees] (FirstName, LastName)
CREATE INDEX IK_Employees_LastName ON [Employees] (LastName)

EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Employees'
  , N'user', N'dbo', N'table', N'Employees'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'(Label: EmployeeID)'
  , N'user', N'dbo', N'table', N'Employees', N'column', N'EmployeeID'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'(Label: Last Name)'
  , N'user', N'dbo', N'table', N'Employees', N'column', N'LastName'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'(Label: First Name)'
  , N'user', N'dbo', N'table', N'Employees', N'column', N'FirstName'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'(Label: Department)'
  , N'user', N'dbo', N'table', N'Employees', N'column', N'DepartmentID'
;

--  Create table 'ItemBranches'
CREATE TABLE [ItemBranches] (
   [ItemID] int NOT NULL
  ,[SubItemID] int NOT NULL
  ,[BranchID] int NOT NULL
  ,[ItemValue] nvarchar(255) NOT NULL
  ,CONSTRAINT PK_ItemBranches PRIMARY KEY NONCLUSTERED (ItemID, SubItemID, BranchID)
)
CREATE INDEX IK_ItemBranches_ItemID_SubItemID ON [ItemBranches] (ItemID, SubItemID)
CREATE INDEX IK_ItemBranches_SubItemID_ItemID ON [ItemBranches] (SubItemID, ItemID)

EXECUTE sp_addextendedproperty N'MS_Description'
  , N''
  , N'user', N'dbo', N'table', N'ItemBranches', N'column', N'ItemID'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N''
  , N'user', N'dbo', N'table', N'ItemBranches', N'column', N'SubItemID'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N''
  , N'user', N'dbo', N'table', N'ItemBranches', N'column', N'BranchID'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N''
  , N'user', N'dbo', N'table', N'ItemBranches', N'column', N'ItemValue'
;

--  Create table 'Items'
CREATE TABLE [Items] (
   [ItemID] int NOT NULL
  ,[SubItemID] int NOT NULL
  ,[ItemName] nvarchar(255) NOT NULL
  ,CONSTRAINT PK_Items PRIMARY KEY CLUSTERED (ItemID, SubItemID)
)

EXECUTE sp_addextendedproperty N'MS_Description'
  , N''
  , N'user', N'dbo', N'table', N'Items', N'column', N'ItemID'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N''
  , N'user', N'dbo', N'table', N'Items', N'column', N'SubItemID'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N''
  , N'user', N'dbo', N'table', N'Items', N'column', N'ItemName'
;

--  Create table 'TestForeignKeyOptions'
CREATE TABLE [TestForeignKeyOptions] (
   [DepartmentID] int NOT NULL
  ,[Memo] nvarchar(50) NOT NULL
  ,CONSTRAINT PK_TestForeignKeyOptions PRIMARY KEY CLUSTERED (DepartmentID)
)

EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Test ForeignKey Actions: CASCADE/DEFAULT/NULL/No Action'
  , N'user', N'dbo', N'table', N'TestForeignKeyOptions'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N''
  , N'user', N'dbo', N'table', N'TestForeignKeyOptions', N'column', N'DepartmentID'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N''
  , N'user', N'dbo', N'table', N'TestForeignKeyOptions', N'column', N'Memo'
;

--  Create table 'TestForeignKeyOptions2'
CREATE TABLE [TestForeignKeyOptions2] (
   [OptionID] int IDENTITY (1,1) NOT NULL
  ,[DepartmentID] int NULL
  ,[Memo] nvarchar(50) NOT NULL
  ,CONSTRAINT PK_TestForeignKeyOptions2 PRIMARY KEY CLUSTERED (OptionID)
)

EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Test ForeignKey Actions: CASCADE/DEFAULT/NULL/No Action'
  , N'user', N'dbo', N'table', N'TestForeignKeyOptions2'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N''
  , N'user', N'dbo', N'table', N'TestForeignKeyOptions2', N'column', N'OptionID'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N''
  , N'user', N'dbo', N'table', N'TestForeignKeyOptions2', N'column', N'DepartmentID'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N''
  , N'user', N'dbo', N'table', N'TestForeignKeyOptions2', N'column', N'Memo'
;

--  Create table 'TestForeignKeyOptions3'
CREATE TABLE [TestForeignKeyOptions3] (
   [OptionID] int IDENTITY (1,1) NOT NULL
  ,[DepartmentID] int NULL
  ,[Memo] nvarchar(50) NOT NULL
  ,CONSTRAINT PK_TestForeignKeyOptions3 PRIMARY KEY CLUSTERED (OptionID)
)

EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Test ForeignKey Actions: CASCADE/DEFAULT/NULL/No Action'
  , N'user', N'dbo', N'table', N'TestForeignKeyOptions3'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N''
  , N'user', N'dbo', N'table', N'TestForeignKeyOptions3', N'column', N'OptionID'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N''
  , N'user', N'dbo', N'table', N'TestForeignKeyOptions3', N'column', N'DepartmentID'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N''
  , N'user', N'dbo', N'table', N'TestForeignKeyOptions3', N'column', N'Memo'
;

--  Create table 'TestForeignKeyOptions4'
CREATE TABLE [TestForeignKeyOptions4] (
   [OptionID] int IDENTITY (1,1) NOT NULL
  ,[DepartmentID] int NULL
  ,[Memo] nvarchar(50) NOT NULL
  ,CONSTRAINT PK_TestForeignKeyOptions4 PRIMARY KEY CLUSTERED (OptionID)
)

EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Test ForeignKey Actions: CASCADE/DEFAULT/NULL/No Action'
  , N'user', N'dbo', N'table', N'TestForeignKeyOptions4'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N''
  , N'user', N'dbo', N'table', N'TestForeignKeyOptions4', N'column', N'OptionID'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N''
  , N'user', N'dbo', N'table', N'TestForeignKeyOptions4', N'column', N'DepartmentID'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N''
  , N'user', N'dbo', N'table', N'TestForeignKeyOptions4', N'column', N'Memo'
;

--  Create table 'ZipCodes'
CREATE TABLE [ZipCodes] (
   [ZipCode] varchar(8) NOT NULL
  ,[Address1] nvarchar(255) NOT NULL
  ,[Address2] nvarchar(255) NOT NULL CONSTRAINT DF_ZipCodes_Address2 DEFAULT ('')
  ,[Address3] nvarchar(255) NOT NULL CONSTRAINT DF_ZipCodes_Address3 DEFAULT ('')
)
CREATE CLUSTERED INDEX IK_ZipCodes_ZipCode ON [ZipCodes] (ZipCode)

EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Zip codes'
  , N'user', N'dbo', N'table', N'ZipCodes'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'Zip code is not unique. (Label: Zip Code)'
  , N'user', N'dbo', N'table', N'ZipCodes', N'column', N'ZipCode'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'(Label: Address1)'
  , N'user', N'dbo', N'table', N'ZipCodes', N'column', N'Address1'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'(Label: Address2)'
  , N'user', N'dbo', N'table', N'ZipCodes', N'column', N'Address2'
EXECUTE sp_addextendedproperty N'MS_Description'
  , N'(Label: Address3)'
  , N'user', N'dbo', N'table', N'ZipCodes', N'column', N'Address3'
;


-- Foreign keys for table 'Departments'
ALTER TABLE [Departments] ADD
  CONSTRAINT FK_Departments_ParentID
  FOREIGN KEY (ParentID)
  REFERENCES Departments(DepartmentID)
, CONSTRAINT FK_Departments_ManagerID
  FOREIGN KEY (ManagerID)
  REFERENCES Employees(EmployeeID)
;
-- Foreign keys for table 'Employees'
ALTER TABLE [Employees] ADD
  CONSTRAINT FK_Employees_DepartmentID
  FOREIGN KEY (DepartmentID)
  REFERENCES Departments(DepartmentID)
;
-- Foreign keys for table 'ItemBranches'
ALTER TABLE [ItemBranches] ADD
  CONSTRAINT FK_ItemBranches_ItemID_SubItemID
  FOREIGN KEY (ItemID,SubItemID)
  REFERENCES Items(ItemID,SubItemID) ON DELETE CASCADE
;
-- Foreign keys for table 'TestForeignKeyOptions'
ALTER TABLE [TestForeignKeyOptions] ADD
  CONSTRAINT FK_TestForeignKeyOptions_DepartmentID
  FOREIGN KEY (DepartmentID)
  REFERENCES Departments(DepartmentID)
;
-- Foreign keys for table 'TestForeignKeyOptions2'
ALTER TABLE [TestForeignKeyOptions2] ADD
  CONSTRAINT FK_TestForeignKeyOptions2_DepartmentID
  FOREIGN KEY (DepartmentID)
  REFERENCES Departments(DepartmentID) ON DELETE CASCADE ON UPDATE SET NULL
;
-- Foreign keys for table 'TestForeignKeyOptions3'
ALTER TABLE [TestForeignKeyOptions3] ADD
  CONSTRAINT FK_TestForeignKeyOptions3_DepartmentID
  FOREIGN KEY (DepartmentID)
  REFERENCES Departments(DepartmentID) ON DELETE SET NULL ON UPDATE SET DEFAULT
;
-- Foreign keys for table 'TestForeignKeyOptions4'
ALTER TABLE [TestForeignKeyOptions4] ADD
  CONSTRAINT FK_TestForeignKeyOptions4_DepartmentID
  FOREIGN KEY (DepartmentID)
  REFERENCES Departments(DepartmentID) ON DELETE SET DEFAULT ON UPDATE CASCADE
;
