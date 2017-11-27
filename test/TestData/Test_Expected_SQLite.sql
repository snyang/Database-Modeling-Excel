DROP TABLE IF EXISTS "categories";

DROP TABLE IF EXISTS "customercustomerdemo";

DROP TABLE IF EXISTS "customerdemographics";

DROP TABLE IF EXISTS "customers";

DROP TABLE IF EXISTS "employeeterritories";

DROP TABLE IF EXISTS "employees";

DROP TABLE IF EXISTS "orderdetails";

DROP TABLE IF EXISTS "orders";

DROP TABLE IF EXISTS "products";

DROP TABLE IF EXISTS "region";

DROP TABLE IF EXISTS "shippers";

DROP TABLE IF EXISTS "suppliers";

DROP TABLE IF EXISTS "territories";

DROP TABLE IF EXISTS "SAMPLETABLE1";

DROP TABLE IF EXISTS "sampletable2";

CREATE TABLE "categories" (
   "CategoryID" integer NOT NULL 
  ,"CategoryName" character varying(15) NOT NULL 
  ,"Description" text NULL 
  ,"Picture" bytea NULL 
  ,CONSTRAINT PK_categories PRIMARY KEY (CategoryID)
);
CREATE INDEX IK_categories_CategoryName ON "categories" (CategoryName);

CREATE TABLE "customercustomerdemo" (
   "CustomerID" character(5) NOT NULL 
  ,"CustomerTypeID" character(10) NOT NULL 
  ,CONSTRAINT PK_customercustomerdemo PRIMARY KEY (CustomerID, CustomerTypeID)
  ,CONSTRAINT IK_customercustomerdemo_CustomerTypeID_CustomerID UNIQUE (CustomerTypeID, CustomerID)
  ,CONSTRAINT FK_customercustomerdemo_CustomerID FOREIGN KEY (CustomerID) REFERENCES customers(CustomerID)
  ,CONSTRAINT FK_customercustomerdemo_CustomerTypeID FOREIGN KEY (CustomerTypeID) REFERENCES customerdemographics(CustomerTypeID)
);

CREATE TABLE "customerdemographics" (
   "CustomerTypeID" character(10) NOT NULL 
  ,"CustomerDesc" text NULL 
  ,CONSTRAINT PK_customerdemographics PRIMARY KEY (CustomerTypeID)
);

CREATE TABLE "customers" (
   "CustomerID" character(5) NOT NULL 
  ,"CompanyName" character varying(40) NOT NULL 
  ,"ContactName" character varying(30) NULL 
  ,"ContactTitle" character varying(30) NULL 
  ,"Address" character varying(60) NULL 
  ,"City" character varying(15) NULL 
  ,"Region" character varying(15) NULL 
  ,"PostalCode" character varying(10) NULL 
  ,"Country" character varying(15) NULL 
  ,"Phone" character varying(24) NULL 
  ,"Fax" character varying(24) NULL 
  ,CONSTRAINT PK_customers PRIMARY KEY (CustomerID)
);
CREATE INDEX IK_customers_City ON "customers" (City);
CREATE INDEX IK_customers_CompanyName ON "customers" (CompanyName);
CREATE INDEX IK_customers_PostalCode ON "customers" (PostalCode);
CREATE INDEX IK_customers_Region ON "customers" (Region);

CREATE TABLE "employeeterritories" (
   "EmployeeID" integer NOT NULL 
  ,"TerritoryID" character varying(20) NOT NULL 
  ,CONSTRAINT PK_employeeterritories PRIMARY KEY (EmployeeID, TerritoryID)
  ,CONSTRAINT FK_employeeterritories_EmployeeID FOREIGN KEY (EmployeeID) REFERENCES employees(EmployeeID)
  ,CONSTRAINT FK_employeeterritories_TerritoryID FOREIGN KEY (TerritoryID) REFERENCES territories(TerritoryID)
);
CREATE INDEX IK_employeeterritories_TerritoryID ON "employeeterritories" (TerritoryID);

CREATE TABLE "employees" (
   "EmployeeID" integer NOT NULL 
  ,"LastName" character varying(20) NOT NULL 
  ,"FirstName" character varying(10) NOT NULL 
  ,"Title" character varying(30) NULL 
  ,"TitleOfCourtesy" character varying(25) NULL 
  ,"BirthDate" date NULL 
  ,"HireDate" date NULL 
  ,"Address" character varying(60) NULL 
  ,"City" character varying(15) NULL 
  ,"Region" character varying(15) NULL 
  ,"PostalCode" character varying(10) NULL 
  ,"Country" character varying(15) NULL 
  ,"HomePhone" character varying(24) NULL 
  ,"Extension" character varying(4) NULL 
  ,"Photo" bytea NULL 
  ,"Notes" text NULL 
  ,"ReportsTo" integer NULL 
  ,"PhotoPath" character varying(255) NULL 
  ,CONSTRAINT PK_employees PRIMARY KEY (EmployeeID)
  ,CONSTRAINT FK_employees_ReportsTo FOREIGN KEY (ReportsTo) REFERENCES employees(EmployeeID)
);
CREATE INDEX IK_employees_LastName ON "employees" (LastName);
CREATE INDEX IK_employees_PostalCode ON "employees" (PostalCode);
CREATE INDEX IK_employees_ReportsTo ON "employees" (ReportsTo);

CREATE TABLE "orderdetails" (
   "OrderID" integer NOT NULL 
  ,"ProductID" integer NOT NULL 
  ,"UnitPrice" numeric NOT NULL 
  ,"Quantity" smallint NOT NULL 
  ,"Discount" numeric NOT NULL 
  ,CONSTRAINT PK_orderdetails PRIMARY KEY (OrderID, ProductID)
  ,CONSTRAINT FK_orderdetails_OrderID FOREIGN KEY (OrderID) REFERENCES orders(OrderID)
  ,CONSTRAINT FK_orderdetails_ProductID FOREIGN KEY (ProductID) REFERENCES products(ProductID)
);
CREATE INDEX IK_orderdetails_OrderID ON "orderdetails" (OrderID);
CREATE INDEX IK_orderdetails_ProductID ON "orderdetails" (ProductID);

CREATE TABLE "orders" (
   "OrderID" integer NOT NULL 
  ,"CustomerID" character(5) NULL 
  ,"EmployeeID" integer NULL 
  ,"OrderDate" date NULL 
  ,"RequiredDate" date NULL 
  ,"ShippedDate" date NULL 
  ,"ShipVia" integer NULL 
  ,"Freight" numeric NULL 
  ,"ShipName" character varying(40) NULL 
  ,"ShipAddress" character varying(60) NULL 
  ,"ShipCity" character varying(15) NULL 
  ,"ShipRegion" character varying(15) NULL 
  ,"ShipPostalCode" character varying(10) NULL 
  ,"ShipCountry" character varying(15) NULL 
  ,CONSTRAINT PK_orders PRIMARY KEY (OrderID)
  ,CONSTRAINT FK_orders_CustomerID FOREIGN KEY (CustomerID) REFERENCES customers(CustomerID)
  ,CONSTRAINT FK_orders_EmployeeID FOREIGN KEY (EmployeeID) REFERENCES employees(EmployeeID)
  ,CONSTRAINT FK_orders_ShipVia FOREIGN KEY (ShipVia) REFERENCES shippers(ShipperID)
);
CREATE INDEX IK_orders_CustomerID ON "orders" (CustomerID);
CREATE INDEX IK_orders_EmployeeID ON "orders" (EmployeeID);
CREATE INDEX IK_orders_OrderDate ON "orders" (OrderDate);
CREATE INDEX IK_orders_ShippedDate ON "orders" (ShippedDate);
CREATE INDEX IK_orders_ShipPostalCode ON "orders" (ShipPostalCode);
CREATE INDEX IK_orders_ShipVia ON "orders" (ShipVia);

CREATE TABLE "products" (
   "ProductID" integer NOT NULL 
  ,"ProductName" character varying(40) NOT NULL 
  ,"SupplierID" integer NULL 
  ,"CategoryID" integer NULL 
  ,"QuantityPerUnit" character varying(20) NULL 
  ,"UnitPrice" numeric NULL 
  ,"UnitsInStock" smallint NULL 
  ,"UnitsOnOrder" smallint NULL 
  ,"ReorderLevel" smallint NULL 
  ,"Discontinued" numeric(1,0) NOT NULL 
  ,CONSTRAINT PK_products PRIMARY KEY (ProductID)
  ,CONSTRAINT FK_products_CategoryID FOREIGN KEY (CategoryID) REFERENCES categories(CategoryID)
  ,CONSTRAINT FK_products_SupplierID FOREIGN KEY (SupplierID) REFERENCES suppliers(SupplierID)
);
CREATE INDEX IK_products_CategoryID ON "products" (CategoryID);
CREATE INDEX IK_products_ProductName ON "products" (ProductName);
CREATE INDEX IK_products_SupplierID ON "products" (SupplierID);

CREATE TABLE "region" (
   "RegionID" integer NOT NULL 
  ,"RegionDescription" character(50) NOT NULL 
  ,CONSTRAINT PK_region PRIMARY KEY (RegionID)
);

CREATE TABLE "shippers" (
   "ShipperID" integer NOT NULL 
  ,"CompanyName" character varying(40) NOT NULL 
  ,"Phone" character varying(24) NULL 
  ,CONSTRAINT PK_shippers PRIMARY KEY (ShipperID)
);

CREATE TABLE "suppliers" (
   "SupplierID" integer NOT NULL 
  ,"CompanyName" character varying(40) NOT NULL 
  ,"ContactName" character varying(30) NULL 
  ,"ContactTitle" character varying(30) NULL 
  ,"Address" character varying(60) NULL 
  ,"City" character varying(15) NULL 
  ,"Region" character varying(15) NULL 
  ,"PostalCode" character varying(10) NULL 
  ,"Country" character varying(15) NULL 
  ,"Phone" character varying(24) NULL 
  ,"Fax" character varying(24) NULL 
  ,"HomePage" text NULL 
  ,CONSTRAINT PK_suppliers PRIMARY KEY (SupplierID)
);
CREATE INDEX IK_suppliers_CompanyName ON "suppliers" (CompanyName);
CREATE INDEX IK_suppliers_PostalCode ON "suppliers" (PostalCode);

CREATE TABLE "territories" (
   "TerritoryID" character varying(20) NOT NULL 
  ,"TerritoryDescription" character(50) NOT NULL 
  ,"RegionID" integer NOT NULL 
  ,CONSTRAINT PK_territories PRIMARY KEY (TerritoryID)
  ,CONSTRAINT FK_territories_RegionID FOREIGN KEY (RegionID) REFERENCES region(RegionID)
);
CREATE INDEX IK_territories_RegionID ON "territories" (RegionID);

CREATE TABLE "SAMPLETABLE1" (
   "ITEMID" numeric NOT NULL 
  ,"SUBITEMID" numeric NOT NULL 
  ,"ITEMNAME" character varying(50) NOT NULL 
  ,CONSTRAINT PK_SAMPLETABLE1 PRIMARY KEY (ITEMID, SUBITEMID)
);

CREATE TABLE "sampletable2" (
   "ItemID" integer NOT NULL 
  ,"SubItemID" integer NOT NULL 
  ,"BranchID" integer NOT NULL 
  ,"ItemName" character varying(50) NOT NULL 
  ,CONSTRAINT PK_sampletable2 PRIMARY KEY (ItemID, SubItemID, BranchID)
  ,CONSTRAINT IK_sampletable2_ItemID_SubItemID UNIQUE (ItemID, SubItemID)
  ,CONSTRAINT FK_sampletable2_ItemID_SubItemID FOREIGN KEY (ItemID,SubItemID) REFERENCES sampletable1(ItemID,SubItemID) ON DELETE CASCADE
);

