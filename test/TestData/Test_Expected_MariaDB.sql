-- Prepare a temporary store procedure to drop foreign keys
DROP PROCEDURE IF EXISTS __tmp_dropFK;
DELIMITER $$
CREATE PROCEDURE __tmp_dropFK (tableName varchar(64))
BEGIN
  DECLARE fkName varchar(64);
  DECLARE sqlDropFK varchar(250);
  DECLARE done INT DEFAULT 0;

  DECLARE fkCursor CURSOR FOR
    SELECT CONSTRAINT_NAME FROM information_schema.TABLE_CONSTRAINTS TC
    WHERE TC.TABLE_SCHEMA = database()
    AND   TC.TABLE_NAME = tableName
    AND   TC.CONSTRAINT_TYPE = 'FOREIGN KEY';
  DECLARE CONTINUE HANDLER FOR SQLSTATE '02000' SET done = 1;

  OPEN fkCursor;
  FETCH fkCursor INTO fkName;
  WHILE done = 0 DO
    SET @sqlDropFK = CONCAT('ALTER TABLE ', tableName ,' DROP FOREIGN KEY ', fkName, ';');
    PREPARE stmt_dropFK FROM @sqlDropFK;
    EXECUTE stmt_dropFK;
    DEALLOCATE PREPARE stmt_dropFK;

    FETCH fkCursor INTO fkName;
  END WHILE;

  CLOSE fkCursor;
END $$
DELIMITER ;

-- Drop foreign keys for table 'AllDataTypes'
CALL __tmp_dropFK('AllDataTypes');

-- Drop foreign keys for table 'Departments'
CALL __tmp_dropFK('Departments');

-- Drop foreign keys for table 'Employees'
CALL __tmp_dropFK('Employees');

-- Drop foreign keys for table 'ItemBranches'
CALL __tmp_dropFK('ItemBranches');

-- Drop foreign keys for table 'Items'
CALL __tmp_dropFK('Items');

-- Drop foreign keys for table 'TestForeignKeyOptions'
CALL __tmp_dropFK('TestForeignKeyOptions');

-- Drop foreign keys for table 'TestForeignKeyOptions2'
CALL __tmp_dropFK('TestForeignKeyOptions2');

-- Drop foreign keys for table 'TestForeignKeyOptions3'
CALL __tmp_dropFK('TestForeignKeyOptions3');

-- Drop foreign keys for table 'ZipCodes'
CALL __tmp_dropFK('ZipCodes');

-- Drop the temporary store procedure of dropping foreign keys
DROP PROCEDURE IF EXISTS __tmp_dropFK;

-- Drop table 'AllDataTypes'
DROP TABLE IF EXISTS AllDataTypes;

-- Drop table 'Departments'
DROP TABLE IF EXISTS Departments;

-- Drop table 'Employees'
DROP TABLE IF EXISTS Employees;

-- Drop table 'ItemBranches'
DROP TABLE IF EXISTS ItemBranches;

-- Drop table 'Items'
DROP TABLE IF EXISTS Items;

-- Drop table 'TestForeignKeyOptions'
DROP TABLE IF EXISTS TestForeignKeyOptions;

-- Drop table 'TestForeignKeyOptions2'
DROP TABLE IF EXISTS TestForeignKeyOptions2;

-- Drop table 'TestForeignKeyOptions3'
DROP TABLE IF EXISTS TestForeignKeyOptions3;

-- Drop table 'ZipCodes'
DROP TABLE IF EXISTS ZipCodes;

-- Create table 'AllDataTypes'
CREATE TABLE AllDataTypes (
   DataTypeID int auto_increment NOT NULL COMMENT '(Label: Data Type ID)'
  ,DataTypeName varchar(15) NOT NULL COMMENT 'Test single quatation (Label: Data Type''s Name)'
  ,DtTinyInt tinyint NOT NULL DEFAULT 1 COMMENT 'Numeric Data Types'
  ,DtTinyIntUnsigned tinyint unsigned NOT NULL DEFAULT 0 COMMENT 'Numeric Data Types'
  ,DtSmallInt smallint NOT NULL DEFAULT 3 COMMENT 'Numeric Data Types'
  ,DtSmallIntUnsigned smallint unsigned NOT NULL DEFAULT 4 COMMENT 'Numeric Data Types'
  ,DtMediumInt mediumint NOT NULL DEFAULT 5 COMMENT 'Numeric Data Types'
  ,DtMediumIntUnsigned mediumint unsigned NOT NULL DEFAULT 6 COMMENT 'Numeric Data Types'
  ,DtInt int NOT NULL DEFAULT 7 COMMENT 'Numeric Data Types'
  ,DtIntUnsigned int unsigned NOT NULL DEFAULT 8 COMMENT 'Numeric Data Types'
  ,DtBigInt bigint NOT NULL DEFAULT 9 COMMENT 'Numeric Data Types'
  ,DtBigIntUnsigned bigint unsigned NOT NULL DEFAULT 10 COMMENT 'Numeric Data Types'
  ,DtDecimal_8_2 decimal(8,2) NOT NULL DEFAULT 11.11 COMMENT 'Numeric Data Types'
  ,DtFloat float NOT NULL COMMENT 'Numeric Data Types'
  ,DtDouble double NOT NULL COMMENT 'Numeric Data Types'
  ,DtBit bit NOT NULL COMMENT 'Numeric Data Types'
  ,DtChar char(255) NOT NULL COMMENT 'String Data Types'
  ,DtNChar char(255) NOT NULL DEFAULT '' COMMENT 'String Data Types'
  ,DtVarchar varchar(255) NOT NULL DEFAULT 'A' COMMENT 'String Data Types'
  ,DtNVarchar varchar(255) NOT NULL DEFAULT 'B' COMMENT 'String Data Types'
  ,DtBinary_8 binary(8) NULL COMMENT 'String Data Types'
  ,DtVarBinary_8 varbinary(8) NULL COMMENT 'String Data Types'
  ,DtTinyBlob tinyblob NULL COMMENT 'String Data Types'
  ,DtBlob blob NULL COMMENT 'String Data Types'
  ,DtMediumBlob mediumblob NULL COMMENT 'String Data Types'
  ,DtLongBlob longblob NULL COMMENT 'String Data Types'
  ,DtTinyText tinytext NULL COMMENT 'String Data Types'
  ,DtText text NULL COMMENT 'String Data Types'
  ,DtMediumText mediumtext NULL COMMENT 'String Data Types'
  ,DtLongText longtext NULL COMMENT 'String Data Types'
  ,DtEnum enum('EnumA','EnumB') NOT NULL DEFAULT 'EnumA' COMMENT 'String Data Types'
  ,DtDate date NOT NULL COMMENT 'Date and Time Data Types'
  ,DtTime time NOT NULL COMMENT 'Date and Time Data Types'
  ,DtDateTime datetime NOT NULL COMMENT 'Date and Time Data Types'
  ,DtYear year NOT NULL COMMENT 'Date and Time Data Types'
  ,DtTimestamp timestamp on update CURRENT_TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'Date and Time Data Types'
  ,DtPoint point NULL COMMENT 'Geometry Types'
  ,DtLineString linestring NULL COMMENT 'Geometry Types'
  ,DtPolygon polygon NULL COMMENT 'Geometry Types'
  ,DtMultiPoint multipoint NULL COMMENT 'Geometry Types'
  ,DtMultiLineString multilinestring NULL COMMENT 'Geometry Types'
  ,DtGeometryCollection geometrycollection NULL COMMENT 'Geometry Types'
  ,DtGeometry geometry NULL COMMENT 'Geometry Types'
  ,CONSTRAINT PK_AllDataTypes PRIMARY KEY (DataTypeID)
  ,CONSTRAINT IK_AllDataTypes_DataTypeName UNIQUE (DataTypeName)
)
   COMMENT = 'Sample table for most common data types'' definitions.'
;

-- Create table 'Departments'
CREATE TABLE Departments (
   DepartmentID int auto_increment NOT NULL COMMENT '(Label: Department ID)'
  ,DepartmentName varchar(50) NOT NULL COMMENT '(Label: Department Name)'
  ,ParentID int NULL COMMENT '(Label: Parent Department)'
  ,ManagerID int NULL COMMENT '(Label: Manager)'
  ,CONSTRAINT PK_Departments PRIMARY KEY (DepartmentID)
)
   COMMENT = 'The department table.'
;

-- Create table 'Employees'
CREATE TABLE Employees (
   EmployeeID int auto_increment NOT NULL COMMENT '(Label: EmployeeID)'
  ,LastName varchar(50) NOT NULL COMMENT '(Label: Last Name)'
  ,FirstName varchar(50) NOT NULL COMMENT '(Label: First Name)'
  ,DepartmentID int NOT NULL COMMENT '(Label: Department)'
  ,CONSTRAINT PK_Employees PRIMARY KEY (EmployeeID)
)
   COMMENT = 'Employees'
;
CREATE INDEX IK_Employees_FirstName_LastName ON Employees (FirstName, LastName);
CREATE INDEX IK_Employees_LastName ON Employees (LastName);

-- Create table 'ItemBranches'
CREATE TABLE ItemBranches (
   ItemID int NOT NULL COMMENT ''
  ,SubItemID int NOT NULL COMMENT ''
  ,BranchID int NOT NULL COMMENT ''
  ,ItemValue varchar(255) NOT NULL COMMENT ''
  ,CONSTRAINT PK_ItemBranches PRIMARY KEY (ItemID, SubItemID, BranchID)
)
   COMMENT = ''
;
CREATE INDEX IK_ItemBranches_ItemID_SubItemID ON ItemBranches (ItemID, SubItemID);
CREATE INDEX IK_ItemBranches_SubItemID_ItemID ON ItemBranches (SubItemID, ItemID);

-- Create table 'Items'
CREATE TABLE Items (
   ItemID int NOT NULL COMMENT ''
  ,SubItemID int NOT NULL COMMENT ''
  ,ItemName varchar(255) NOT NULL COMMENT ''
  ,CONSTRAINT PK_Items PRIMARY KEY (ItemID, SubItemID)
)
   COMMENT = ''
;

-- Create table 'TestForeignKeyOptions'
CREATE TABLE TestForeignKeyOptions (
   DepartmentID int NOT NULL COMMENT ''
  ,Memo varchar(50) NOT NULL COMMENT ''
  ,CONSTRAINT PK_TestForeignKeyOptions PRIMARY KEY (DepartmentID)
)
   COMMENT = 'Test ForeignKey Actions: CASCADE/NULL/No Action'
;

-- Create table 'TestForeignKeyOptions2'
CREATE TABLE TestForeignKeyOptions2 (
   OptionID int auto_increment NOT NULL COMMENT ''
  ,DepartmentID int NULL COMMENT ''
  ,Memo varchar(50) NOT NULL COMMENT ''
  ,CONSTRAINT PK_TestForeignKeyOptions2 PRIMARY KEY (OptionID)
)
   COMMENT = 'Test ForeignKey Actions: CASCADE/NULL/No Action'
;

-- Create table 'TestForeignKeyOptions3'
CREATE TABLE TestForeignKeyOptions3 (
   OptionID int auto_increment NOT NULL COMMENT ''
  ,DepartmentID int NULL COMMENT ''
  ,Memo varchar(50) NOT NULL COMMENT ''
  ,CONSTRAINT PK_TestForeignKeyOptions3 PRIMARY KEY (OptionID)
)
   COMMENT = 'Test ForeignKey Actions: CASCADE/NULL/No Action'
;

-- Create table 'ZipCodes'
CREATE TABLE ZipCodes (
   ZipCode varchar(8) NOT NULL COMMENT 'Zip code is not unique. (Label: Zip Code)'
  ,Address1 varchar(255) NOT NULL COMMENT '(Label: Address1)'
  ,Address2 varchar(255) NOT NULL DEFAULT '' COMMENT '(Label: Address2)'
  ,Address3 varchar(255) NOT NULL DEFAULT '' COMMENT '(Label: Address3)'
)
   COMMENT = 'Zip codes'
;
CREATE INDEX IK_ZipCodes_ZipCode ON ZipCodes (ZipCode);


-- Foreign keys for table 'Departments'
ALTER TABLE Departments
  ADD CONSTRAINT FK_Departments_ManagerID
  FOREIGN KEY (ManagerID)
  REFERENCES Employees(EmployeeID);
ALTER TABLE Departments
  ADD CONSTRAINT FK_Departments_ParentID
  FOREIGN KEY (ParentID)
  REFERENCES Departments(DepartmentID);

-- Foreign keys for table 'Employees'
ALTER TABLE Employees
  ADD CONSTRAINT FK_Employees_DepartmentID
  FOREIGN KEY (DepartmentID)
  REFERENCES Departments(DepartmentID);

-- Foreign keys for table 'ItemBranches'
ALTER TABLE ItemBranches
  ADD CONSTRAINT FK_ItemBranches_ItemID_SubItemID
  FOREIGN KEY (ItemID,SubItemID)
  REFERENCES Items(ItemID,SubItemID) ON DELETE CASCADE;

-- Foreign keys for table 'TestForeignKeyOptions'
ALTER TABLE TestForeignKeyOptions
  ADD CONSTRAINT FK_TestForeignKeyOptions_DepartmentID
  FOREIGN KEY (DepartmentID)
  REFERENCES Departments(DepartmentID);

-- Foreign keys for table 'TestForeignKeyOptions2'
ALTER TABLE TestForeignKeyOptions2
  ADD CONSTRAINT FK_TestForeignKeyOptions2_DepartmentID
  FOREIGN KEY (DepartmentID)
  REFERENCES Departments(DepartmentID) ON DELETE CASCADE ON UPDATE SET NULL;

-- Foreign keys for table 'TestForeignKeyOptions3'
ALTER TABLE TestForeignKeyOptions3
  ADD CONSTRAINT FK_TestForeignKeyOptions3_DepartmentID
  FOREIGN KEY (DepartmentID)
  REFERENCES Departments(DepartmentID) ON DELETE SET NULL ON UPDATE CASCADE;

