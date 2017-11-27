DROP TABLE IF EXISTS categories CASCADE;

DROP TABLE IF EXISTS customerdemo CASCADE;

DROP TABLE IF EXISTS customerdemographics CASCADE;

DROP TABLE IF EXISTS customers CASCADE;

DROP TABLE IF EXISTS employees CASCADE;

DROP TABLE IF EXISTS employeeterritories CASCADE;

DROP TABLE IF EXISTS orderdetails CASCADE;

DROP TABLE IF EXISTS orders CASCADE;

DROP TABLE IF EXISTS products CASCADE;

DROP TABLE IF EXISTS region CASCADE;

DROP TABLE IF EXISTS shippers CASCADE;

DROP TABLE IF EXISTS suppliers CASCADE;

DROP TABLE IF EXISTS territories CASCADE;

DROP TABLE IF EXISTS sampletable1 CASCADE;

DROP TABLE IF EXISTS sampletable2 CASCADE;

CREATE TABLE categories (
   categoryid int4 NOT NULL 
  ,categoryname varchar(15) NOT NULL 
  ,description text NULL 
  ,picture bytea NULL 
  ,CONSTRAINT PK_categories PRIMARY KEY (categoryid)
);
COMMENT ON TABLE categories IS 'categories';
CREATE INDEX I_categories_categoryname ON categories (categoryname);

CREATE TABLE customerdemo (
   customerid bpchar NOT NULL 
  ,customertypeid bpchar NOT NULL 
  ,CONSTRAINT PK_customerdemo PRIMARY KEY (customerid, customertypeid)
  ,CONSTRAINT UI_customerdemo_customertypeid_customerid UNIQUE (customertypeid, customerid)
);
COMMENT ON TABLE customerdemo IS 'customerdemo';

CREATE TABLE customerdemographics (
   customertypeid bpchar NOT NULL 
  ,customerdesc text NULL 
  ,CONSTRAINT PK_customerdemographics PRIMARY KEY (customertypeid)
);
COMMENT ON TABLE customerdemographics IS 'customerdemographics';

CREATE TABLE customers (
   customerid bpchar NOT NULL 
  ,companyname varchar(40) NOT NULL 
  ,contactname varchar(30) NULL 
  ,contacttitle varchar(30) NULL 
  ,address varchar(60) NULL 
  ,city varchar(15) NULL 
  ,region varchar(15) NULL 
  ,postalcode varchar(10) NULL 
  ,country varchar(15) NULL 
  ,phone varchar(24) NULL 
  ,fax varchar(24) NULL 
  ,CONSTRAINT PK_customers PRIMARY KEY (customerid)
);
COMMENT ON TABLE customers IS 'customers';
CREATE INDEX I_customers_city ON customers (city);
CREATE INDEX I_customers_companyname ON customers (companyname);
CREATE INDEX I_customers_postalcode ON customers (postalcode);
CREATE INDEX I_customers_region ON customers (region);

CREATE TABLE employees (
   employeeid int4 NOT NULL 
  ,lastname varchar(20) NOT NULL 
  ,firstname varchar(10) NOT NULL 
  ,title varchar(30) NULL 
  ,titleofcourtesy varchar(25) NULL 
  ,birthdate date NULL 
  ,hiredate date NULL 
  ,address varchar(60) NULL 
  ,city varchar(15) NULL 
  ,region varchar(15) NULL 
  ,postalcode varchar(10) NULL 
  ,country varchar(15) NULL 
  ,homephone varchar(24) NULL 
  ,extension varchar(4) NULL 
  ,photo bytea NULL 
  ,notes text NULL 
  ,reportsto int4 NULL 
  ,photopath varchar(255) NULL 
  ,CONSTRAINT PK_employees PRIMARY KEY (employeeid)
);
COMMENT ON TABLE employees IS 'employees';
CREATE INDEX I_employees_lastname ON employees (lastname);
CREATE INDEX I_employees_postalcode ON employees (postalcode);
CREATE INDEX I_employees_reportsto ON employees (reportsto);

CREATE TABLE employeeterritories (
   employeeid int4 NOT NULL 
  ,territoryid varchar(20) NOT NULL 
  ,CONSTRAINT PK_employeeterritories PRIMARY KEY (employeeid, territoryid)
);
COMMENT ON TABLE employeeterritories IS 'employeeterritories';
CREATE INDEX I_employeeterritories_territoryid ON employeeterritories (territoryid);

CREATE TABLE orderdetails (
   orderid int4 NOT NULL 
  ,productid int4 NOT NULL 
  ,unitprice numeric NOT NULL 
  ,quantity int2 NOT NULL DEFAULT 100
  ,discount numeric(12,4) NOT NULL DEFAULT 100
  ,CONSTRAINT PK_orderdetails PRIMARY KEY (orderid, productid)
);
COMMENT ON TABLE orderdetails IS 'orderdetails';
CREATE INDEX I_orderdetails_orderid ON orderdetails (orderid);
CREATE INDEX I_orderdetails_productid ON orderdetails (productid);

CREATE TABLE orders (
   orderid int4 NOT NULL 
  ,customerid bpchar NULL 
  ,employeeid int4 NULL 
  ,orderdate date NULL 
  ,requireddate date NULL 
  ,shippeddate date NULL 
  ,shipvia int4 NULL 
  ,freight numeric NULL 
  ,shipname varchar(40) NULL 
  ,shipaddress varchar(60) NULL 
  ,shipcity varchar(15) NULL 
  ,shipregion varchar(15) NULL 
  ,shippostalcode varchar(10) NULL 
  ,shipcountry varchar(15) NULL 
  ,CONSTRAINT PK_orders PRIMARY KEY (orderid)
);
COMMENT ON TABLE orders IS 'orders';
CREATE INDEX I_orders_customerid ON orders (customerid);
CREATE INDEX I_orders_employeeid ON orders (employeeid);
CREATE INDEX I_orders_orderdate ON orders (orderdate);
CREATE INDEX I_orders_shippeddate ON orders (shippeddate);
CREATE INDEX I_orders_shippostalcode ON orders (shippostalcode);
CREATE INDEX I_orders_shipvia ON orders (shipvia);

CREATE TABLE products (
   productid int4 NOT NULL 
  ,productname varchar(40) NOT NULL 
  ,supplierid int4 NULL 
  ,categoryid int4 NULL 
  ,quantityperunit varchar(20) NULL 
  ,unitprice numeric NULL 
  ,unitsinstock int2 NULL 
  ,unitsonorder int2 NULL 
  ,reorderlevel int2 NULL 
  ,discontinued numeric(1,0) NOT NULL 
  ,CONSTRAINT PK_products PRIMARY KEY (productid)
);
COMMENT ON TABLE products IS 'products';
CREATE INDEX I_products_categoryid ON products (categoryid);
CREATE INDEX I_products_productname ON products (productname);
CREATE INDEX I_products_supplierid ON products (supplierid);

CREATE TABLE region (
   regionid int4 NOT NULL 
  ,regiondescription bpchar NOT NULL 
  ,CONSTRAINT PK_region PRIMARY KEY (regionid)
);
COMMENT ON TABLE region IS 'region';

CREATE TABLE shippers (
   shipperid int4 NOT NULL 
  ,companyname varchar(40) NOT NULL 
  ,phone varchar(24) NULL 
  ,CONSTRAINT PK_shippers PRIMARY KEY (shipperid)
);
COMMENT ON TABLE shippers IS 'shippers';

CREATE TABLE suppliers (
   supplierid int4 NOT NULL 
  ,companyname varchar(40) NOT NULL 
  ,contactname varchar(30) NULL 
  ,contacttitle varchar(30) NULL 
  ,address varchar(60) NULL 
  ,city varchar(15) NULL 
  ,region varchar(15) NULL 
  ,postalcode varchar(10) NULL 
  ,country varchar(15) NULL 
  ,phone varchar(24) NULL 
  ,fax varchar(24) NULL 
  ,homepage text NULL 
  ,CONSTRAINT PK_suppliers PRIMARY KEY (supplierid)
);
COMMENT ON TABLE suppliers IS 'suppliers';
CREATE INDEX I_suppliers_companyname ON suppliers (companyname);
CREATE INDEX I_suppliers_postalcode ON suppliers (postalcode);

CREATE TABLE territories (
   territoryid varchar(20) NOT NULL 
  ,territorydescription bpchar NOT NULL 
  ,regionid int4 NOT NULL 
  ,CONSTRAINT PK_territories PRIMARY KEY (territoryid)
);
COMMENT ON TABLE territories IS 'territories';
CREATE INDEX I_territories_regionid ON territories (regionid);

CREATE TABLE sampletable1 (
   itemid numeric NOT NULL 
  ,subitemid numeric NOT NULL 
  ,itemname varchar(50) NOT NULL 
  ,CONSTRAINT PK_sampletable1 PRIMARY KEY (itemid, subitemid)
);
COMMENT ON TABLE sampletable1 IS 'sampletable1';

CREATE TABLE sampletable2 (
   itemid int4 NOT NULL 
  ,subitemid int4 NOT NULL 
  ,branchid int4 NOT NULL 
  ,itemname varchar(50) NOT NULL 
  ,CONSTRAINT PK_sampletable2 PRIMARY KEY (itemid, subitemid, branchid)
  ,CONSTRAINT UI_sampletable2_itemid_subitemid UNIQUE (itemid, subitemid)
  ,CONSTRAINT UI_sampletable2_subitemid_itemid UNIQUE (subitemid, itemid)
);
COMMENT ON TABLE sampletable2 IS 'sampletable2';


ALTER TABLE customerdemo ADD CONSTRAINT FK_customerdemo_customerid FOREIGN KEY (customerid) REFERENCES customers(customerid);
ALTER TABLE customerdemo ADD CONSTRAINT FK_customerdemo_customertypeid FOREIGN KEY (customertypeid) REFERENCES customerdemographics(customertypeid);

ALTER TABLE employees ADD CONSTRAINT FK_employees_reportsto FOREIGN KEY (reportsto) REFERENCES employees(employeeid);

ALTER TABLE employeeterritories ADD CONSTRAINT FK_employeeterritories_employeeid FOREIGN KEY (employeeid) REFERENCES employees(employeeid);
ALTER TABLE employeeterritories ADD CONSTRAINT FK_employeeterritories_territoryid FOREIGN KEY (territoryid) REFERENCES territories(territoryid);

ALTER TABLE orderdetails ADD CONSTRAINT FK_orderdetails_orderid FOREIGN KEY (orderid) REFERENCES orders(orderid);
ALTER TABLE orderdetails ADD CONSTRAINT FK_orderdetails_productid FOREIGN KEY (productid) REFERENCES products(productid);

ALTER TABLE orders ADD CONSTRAINT FK_orders_customerid FOREIGN KEY (customerid) REFERENCES customers(customerid);
ALTER TABLE orders ADD CONSTRAINT FK_orders_employeeid FOREIGN KEY (employeeid) REFERENCES employees(employeeid);
ALTER TABLE orders ADD CONSTRAINT FK_orders_shipvia FOREIGN KEY (shipvia) REFERENCES shippers(shipperid);

ALTER TABLE products ADD CONSTRAINT FK_products_categoryid FOREIGN KEY (categoryid) REFERENCES categories(categoryid);
ALTER TABLE products ADD CONSTRAINT FK_products_categoryid FOREIGN KEY (categoryid) REFERENCES suppliers(categoryid);

ALTER TABLE territories ADD CONSTRAINT FK_territories_categoryid FOREIGN KEY (categoryid) REFERENCES region(categoryid);

ALTER TABLE sampletable2 ADD CONSTRAINT FK_sampletable2_itemid_subitemid FOREIGN KEY (itemid,subitemid) REFERENCES sampletable1(itemid,subitemid) ON DELETE CASCADE;



COMMENT ON COLUMN categories.categoryid IS '(Label: categoryid)';
COMMENT ON COLUMN categories.categoryname IS '(Label: categoryname)';
COMMENT ON COLUMN categories.description IS '(Label: description)';
COMMENT ON COLUMN categories.picture IS '(Label: picture)';


COMMENT ON COLUMN customerdemo.customerid IS '(Label: customerid)';
COMMENT ON COLUMN customerdemo.customertypeid IS '(Label: customertypeid)';


COMMENT ON COLUMN customerdemographics.customertypeid IS '(Label: customertypeid)';
COMMENT ON COLUMN customerdemographics.customerdesc IS '(Label: customerdesc)';


COMMENT ON COLUMN customers.customerid IS '(Label: customerid)';
COMMENT ON COLUMN customers.companyname IS '(Label: companyname)';
COMMENT ON COLUMN customers.contactname IS '(Label: contactname)';
COMMENT ON COLUMN customers.contacttitle IS '(Label: contacttitle)';
COMMENT ON COLUMN customers.address IS '(Label: address)';
COMMENT ON COLUMN customers.city IS '(Label: city)';
COMMENT ON COLUMN customers.region IS '(Label: region)';
COMMENT ON COLUMN customers.postalcode IS '(Label: postalcode)';
COMMENT ON COLUMN customers.country IS '(Label: country)';
COMMENT ON COLUMN customers.phone IS '(Label: phone)';
COMMENT ON COLUMN customers.fax IS '(Label: fax)';


COMMENT ON COLUMN employees.employeeid IS '(Label: employeeid)';
COMMENT ON COLUMN employees.lastname IS '(Label: lastname)';
COMMENT ON COLUMN employees.firstname IS '(Label: firstname)';
COMMENT ON COLUMN employees.title IS '(Label: title)';
COMMENT ON COLUMN employees.titleofcourtesy IS '(Label: titleofcourtesy)';
COMMENT ON COLUMN employees.birthdate IS '(Label: birthdate)';
COMMENT ON COLUMN employees.hiredate IS '(Label: hiredate)';
COMMENT ON COLUMN employees.address IS '(Label: address)';
COMMENT ON COLUMN employees.city IS '(Label: city)';
COMMENT ON COLUMN employees.region IS '(Label: region)';
COMMENT ON COLUMN employees.postalcode IS '(Label: postalcode)';
COMMENT ON COLUMN employees.country IS '(Label: country)';
COMMENT ON COLUMN employees.homephone IS '(Label: homephone)';
COMMENT ON COLUMN employees.extension IS '(Label: extension)';
COMMENT ON COLUMN employees.photo IS '(Label: photo)';
COMMENT ON COLUMN employees.notes IS '(Label: notes)';
COMMENT ON COLUMN employees.reportsto IS '(Label: reportsto)';
COMMENT ON COLUMN employees.photopath IS '(Label: photopath)';


COMMENT ON COLUMN employeeterritories.employeeid IS '(Label: employeeid)';
COMMENT ON COLUMN employeeterritories.territoryid IS '(Label: territoryid)';


COMMENT ON COLUMN orderdetails.orderid IS '(Label: orderid)';
COMMENT ON COLUMN orderdetails.productid IS '(Label: productid)';
COMMENT ON COLUMN orderdetails.unitprice IS '(Label: unitprice)';
COMMENT ON COLUMN orderdetails.quantity IS '(Label: quantity)';
COMMENT ON COLUMN orderdetails.discount IS '(Label: discount)';


COMMENT ON COLUMN orders.orderid IS '(Label: orderid)';
COMMENT ON COLUMN orders.customerid IS '(Label: customerid)';
COMMENT ON COLUMN orders.employeeid IS '(Label: employeeid)';
COMMENT ON COLUMN orders.orderdate IS '(Label: orderdate)';
COMMENT ON COLUMN orders.requireddate IS '(Label: requireddate)';
COMMENT ON COLUMN orders.shippeddate IS '(Label: shippeddate)';
COMMENT ON COLUMN orders.shipvia IS '(Label: shipvia)';
COMMENT ON COLUMN orders.freight IS '(Label: freight)';
COMMENT ON COLUMN orders.shipname IS '(Label: shipname)';
COMMENT ON COLUMN orders.shipaddress IS '(Label: shipaddress)';
COMMENT ON COLUMN orders.shipcity IS '(Label: shipcity)';
COMMENT ON COLUMN orders.shipregion IS '(Label: shipregion)';
COMMENT ON COLUMN orders.shippostalcode IS '(Label: shippostalcode)';
COMMENT ON COLUMN orders.shipcountry IS '(Label: shipcountry)';


COMMENT ON COLUMN products.productid IS '(Label: productid)';
COMMENT ON COLUMN products.productname IS '(Label: productname)';
COMMENT ON COLUMN products.supplierid IS '(Label: supplierid)';
COMMENT ON COLUMN products.categoryid IS '(Label: categoryid)';
COMMENT ON COLUMN products.quantityperunit IS '(Label: quantityperunit)';
COMMENT ON COLUMN products.unitprice IS '(Label: unitprice)';
COMMENT ON COLUMN products.unitsinstock IS '(Label: unitsinstock)';
COMMENT ON COLUMN products.unitsonorder IS '(Label: unitsonorder)';
COMMENT ON COLUMN products.reorderlevel IS '(Label: reorderlevel)';
COMMENT ON COLUMN products.discontinued IS '(Label: discontinued)';


COMMENT ON COLUMN region.regionid IS '(Label: regionid)';
COMMENT ON COLUMN region.regiondescription IS '(Label: regiondescription)';


COMMENT ON COLUMN shippers.shipperid IS '(Label: shipperid)';
COMMENT ON COLUMN shippers.companyname IS '(Label: companyname)';
COMMENT ON COLUMN shippers.phone IS '(Label: phone)';


COMMENT ON COLUMN suppliers.supplierid IS '(Label: supplierid)';
COMMENT ON COLUMN suppliers.companyname IS '(Label: companyname)';
COMMENT ON COLUMN suppliers.contactname IS '(Label: contactname)';
COMMENT ON COLUMN suppliers.contacttitle IS '(Label: contacttitle)';
COMMENT ON COLUMN suppliers.address IS '(Label: address)';
COMMENT ON COLUMN suppliers.city IS '(Label: city)';
COMMENT ON COLUMN suppliers.region IS '(Label: region)';
COMMENT ON COLUMN suppliers.postalcode IS '(Label: postalcode)';
COMMENT ON COLUMN suppliers.country IS '(Label: country)';
COMMENT ON COLUMN suppliers.phone IS '(Label: phone)';
COMMENT ON COLUMN suppliers.fax IS '(Label: fax)';
COMMENT ON COLUMN suppliers.homepage IS '(Label: homepage)';


COMMENT ON COLUMN territories.territoryid IS '(Label: territoryid)';
COMMENT ON COLUMN territories.territorydescription IS '(Label: territorydescription)';
COMMENT ON COLUMN territories.regionid IS '(Label: regionid)';


COMMENT ON COLUMN sampletable1.itemid IS '(Label: itemid)';
COMMENT ON COLUMN sampletable1.subitemid IS '(Label: subitemid)';
COMMENT ON COLUMN sampletable1.itemname IS '(Label: itemname)';


COMMENT ON COLUMN sampletable2.itemid IS '(Label: itemid)';
COMMENT ON COLUMN sampletable2.subitemid IS '(Label: subitemid)';
COMMENT ON COLUMN sampletable2.branchid IS '(Label: branchid)';
COMMENT ON COLUMN sampletable2.itemname IS '(Label: itemname)';

