-- ----------------------------------------------------------------------
-- MySQL Migration Toolkit
-- SQL Create Script
-- ----------------------------------------------------------------------

SET FOREIGN_KEY_CHECKS = 0;

CREATE DATABASE IF NOT EXISTS `db2`
  CHARACTER SET latin1 COLLATE latin1_swedish_ci;
USE `db2`;
-- -------------------------------------
-- Tables

DROP TABLE IF EXISTS `db2`.`Account`;
CREATE TABLE `db2`.`Account` (
  `ID` INT(10) NOT NULL AUTO_INCREMENT,
  `Admin` VARCHAR(50) NULL,
  `Password` VARCHAR(50) NULL,
  PRIMARY KEY (`ID`),
  UNIQUE INDEX `Admin` (`Admin`),
  INDEX `ID` (`ID`)
)
ENGINE = INNODB;

DROP TABLE IF EXISTS `db2`.`AdminUser`;
CREATE TABLE `db2`.`AdminUser` (
  `ID` INT(10) NOT NULL AUTO_INCREMENT,
  `Admin` VARCHAR(50) NULL,
  `Password` VARCHAR(50) NULL,
  PRIMARY KEY (`ID`),
  UNIQUE INDEX `Admin` (`Admin`),
  UNIQUE INDEX `Password` (`Password`),
  INDEX `ID` (`ID`)
)
ENGINE = INNODB;

DROP TABLE IF EXISTS `db2`.`Categories`;
CREATE TABLE `db2`.`Categories` (
  `CategoryID` INT(10) NOT NULL AUTO_INCREMENT,
  `CategoryName` VARCHAR(50) NULL,
  `CategoryDecsription` LONGTEXT NULL,
  `FileName` VARCHAR(255) NULL,
  `CategoryPhoto` LONGBLOB NULL,
  PRIMARY KEY (`CategoryID`),
  INDEX `CategoryName` (`CategoryName`)
)
ENGINE = INNODB;

DROP TABLE IF EXISTS `db2`.`CompanySetup`;
CREATE TABLE `db2`.`CompanySetup` (
  `SetupID` INT(10) NOT NULL AUTO_INCREMENT,
  `SalesTaxRate` DOUBLE(15, 5) NULL,
  `CompanyName` VARCHAR(50) NULL,
  `Address` LONGTEXT NULL,
  `City` VARCHAR(50) NULL,
  `StateOrProvince` VARCHAR(20) NULL,
  `PostalCode` VARCHAR(20) NULL,
  `Country` VARCHAR(50) NULL,
  `PhoneNumber` VARCHAR(30) NULL,
  `FaxNumber` VARCHAR(30) NULL,
  `DefaultPaymentTerms` VARCHAR(255) NULL,
  `DefaultInvoiceDescription` LONGTEXT NULL,
  `Logo` LONGBLOB NULL,
  `FileName` VARCHAR(255) NULL,
  PRIMARY KEY (`SetupID`),
  INDEX `CompanyName` (`CompanyName`),
  INDEX `My Company InformationSalesTaxR` (`SalesTaxRate`),
  INDEX `PostalCode` (`PostalCode`)
)
ENGINE = INNODB;

DROP TABLE IF EXISTS `db2`.`Customers`;
CREATE TABLE `db2`.`Customers` (
  `CustomerID` INT(10) NOT NULL AUTO_INCREMENT,
  `ContactFirstName` VARCHAR(30) NULL,
  `ContactLastName` VARCHAR(50) NULL,
  `CompanyName` VARCHAR(50) NULL,
  `BillingAddress` VARCHAR(255) NULL,
  `City` VARCHAR(50) NULL,
  `StateOrProvince` VARCHAR(20) NULL,
  `PostalCode` VARCHAR(50) NULL,
  `Country` VARCHAR(50) NULL,
  `ContactTitle` VARCHAR(50) NULL,
  `PhoneNumber` VARCHAR(30) NULL,
  `FaxNumber` VARCHAR(30) NULL,
  `AccountNum` VARCHAR(50) NULL,
  PRIMARY KEY (`CustomerID`),
  UNIQUE INDEX `AccountNum` (`AccountNum`),
  UNIQUE INDEX `CompanyName` (`CompanyName`),
  INDEX `ContactLastName` (`ContactLastName`)
)
ENGINE = INNODB;

DROP TABLE IF EXISTS `db2`.`Employees`;
CREATE TABLE `db2`.`Employees` (
  `EmployeeID` INT(10) NOT NULL AUTO_INCREMENT,
  `FullName` VARCHAR(50) NULL,
  `Title` VARCHAR(50) NULL,
  `Address` VARCHAR(50) NULL,
  `City` VARCHAR(50) NULL,
  `State` VARCHAR(50) NULL,
  `ZipCode` INT(10) NULL,
  `SSNNumber` INT(10) NULL,
  `HireDate` DATETIME NULL,
  `ContactPhone` VARCHAR(30) NULL,
  `BillingRate` DECIMAL(19, 4) NULL,
  `Photo` LONGBLOB NULL,
  `FileName` VARCHAR(255) NULL,
  `Notes` LONGTEXT NULL,
  `Total Hours` INT(10) NULL,
  `OTHours` INT(10) NULL,
  PRIMARY KEY (`EmployeeID`),
  UNIQUE INDEX `SSNNumber` (`SSNNumber`),
  INDEX `ZipCode` (`ZipCode`)
)
ENGINE = INNODB;

DROP TABLE IF EXISTS `db2`.`EmpsTimeTable`;
CREATE TABLE `db2`.`EmpsTimeTable` (
  `EmployeeID` INT(10) NOT NULL,
  `Mon` INT(10) NULL,
  `Tues` INT(10) NULL,
  `Wed` INT(10) NULL,
  `Thurs` INT(10) NULL,
  `Fri` INT(10) NULL,
  `Sat` INT(10) NULL,
  `Sun` INT(10) NULL,
  PRIMARY KEY (`EmployeeID`),
  UNIQUE INDEX `{A4BEA8EF-D494-4F4E-8354-2B10F8` (`EmployeeID`),
  INDEX `EmployeeID` (`EmployeeID`)
)
ENGINE = INNODB;

DROP TABLE IF EXISTS `db2`.`Parts`;
CREATE TABLE `db2`.`Parts` (
  `PartID` INT(10) NOT NULL AUTO_INCREMENT,
  `CategoryID` INT(10) NULL,
  `PartName` VARCHAR(50) NULL,
  `PartDescription` VARCHAR(255) NULL,
  `UnitPrice` DECIMAL(19, 4) NULL,
  `UnitsStock` INT(10) NULL,
  `UnitsOnOrder` INT(10) NULL,
  `ReorderLevel` INT(10) NULL,
  `FileName` VARCHAR(255) NULL,
  `PartImage` LONGBLOB NULL,
  `PartCode` VARCHAR(50) NULL,
  PRIMARY KEY (`PartID`),
  UNIQUE INDEX `partcode` (`PartCode`),
  INDEX `CategoryID` (`CategoryID`)
)
ENGINE = INNODB;

DROP TABLE IF EXISTS `db2`.`Payment Methods`;
CREATE TABLE `db2`.`Payment Methods` (
  `PaymentMethodID` INT(10) NOT NULL AUTO_INCREMENT,
  `PaymentMethod` VARCHAR(50) NULL,
  `CreditCard` TINYINT(1) NOT NULL,
  PRIMARY KEY (`PaymentMethodID`)
)
ENGINE = INNODB;

DROP TABLE IF EXISTS `db2`.`Payments`;
CREATE TABLE `db2`.`Payments` (
  `PaymentID` INT(10) NOT NULL AUTO_INCREMENT,
  `WorkorderID` INT(10) NULL,
  `PaymentMethodID` INT(10) NULL,
  `CustomerName` VARCHAR(50) NULL,
  `PaymentTime` DATETIME NULL,
  `PaymentDate` DATETIME NULL,
  `PaymentAmount` DECIMAL(19, 4) NULL,
  `CheckNumber` SMALLINT(5) NULL,
  `CreditCardNumber` VARCHAR(30) NULL,
  `CardholdersName` VARCHAR(50) NULL,
  `CreditCardExpDate` DATETIME NULL,
  `CreditCardAuthorizationNumber` VARCHAR(30) NULL,
  PRIMARY KEY (`PaymentID`),
  INDEX `{B73F9D58-FE1E-4B66-8676-239BCB` (`WorkorderID`),
  INDEX `PaymentMethodID` (`PaymentMethodID`),
  INDEX `WorkorderID` (`WorkorderID`)
)
ENGINE = INNODB;

DROP TABLE IF EXISTS `db2`.`ServerIP`;
CREATE TABLE `db2`.`ServerIP` (
  `ID` INT(10) NOT NULL AUTO_INCREMENT,
  `ServerIP` VARCHAR(50) NULL,
  PRIMARY KEY (`ID`),
  INDEX `ID` (`ID`)
)
ENGINE = INNODB;

DROP TABLE IF EXISTS `db2`.`States`;
CREATE TABLE `db2`.`States` (
  `ID` INT(10) NOT NULL AUTO_INCREMENT,
  `States` VARCHAR(50) NULL,
  PRIMARY KEY (`ID`)
)
ENGINE = INNODB;

DROP TABLE IF EXISTS `db2`.`StockTrack`;
CREATE TABLE `db2`.`StockTrack` (
  `ID` INT(10) NOT NULL AUTO_INCREMENT,
  `PartID` VARCHAR(50) NULL,
  `PartCode` VARCHAR(50) NULL,
  `PartName` VARCHAR(50) NULL,
  `PartDescription` VARCHAR(50) NULL,
  `StockDate` DATETIME NULL,
  `ReceivedAmount` VARCHAR(50) NULL,
  `Vendor` VARCHAR(50) NULL,
  `VendorAcctNum` VARCHAR(50) NULL,
  PRIMARY KEY (`ID`),
  INDEX `ID` (`ID`),
  INDEX `PartID` (`PartID`)
)
ENGINE = INNODB;

DROP TABLE IF EXISTS `db2`.`Tax`;
CREATE TABLE `db2`.`Tax` (
  `EmpID` INT(10) NOT NULL,
  `fedtax` DECIMAL(19, 4) NULL,
  `statetax` DECIMAL(19, 4) NULL,
  `fica` DECIMAL(19, 4) NULL,
  `socialsec` DECIMAL(19, 4) NULL,
  `advance` DECIMAL(19, 4) NULL,
  `garn` DECIMAL(19, 4) NULL,
  `childsupport` DECIMAL(19, 4) NULL,
  `user_defined` DECIMAL(19, 4) NULL,
  `user_defined1` DECIMAL(19, 4) NULL,
  `user_defined2` DECIMAL(19, 4) NULL,
  `exemptions` INT(10) NULL,
  PRIMARY KEY (`EmpID`),
  UNIQUE INDEX `{E6667C12-50B3-4E07-BFC2-C7787F` (`EmpID`),
  INDEX `EmpID` (`EmpID`)
)
ENGINE = INNODB;

DROP TABLE IF EXISTS `db2`.`Users`;
CREATE TABLE `db2`.`Users` (
  `LoginID` VARCHAR(50) NOT NULL,
  `User` VARCHAR(50) NULL,
  `Password` VARCHAR(50) NULL,
  `IP` VARCHAR(50) NULL,
  PRIMARY KEY (`LoginID`),
  UNIQUE INDEX `LoginID` (`LoginID`)
)
ENGINE = INNODB;

DROP TABLE IF EXISTS `db2`.`Vendors`;
CREATE TABLE `db2`.`Vendors` (
  `VenderID` INT(10) NOT NULL AUTO_INCREMENT,
  `AddDate` VARCHAR(50) NULL,
  `AccountNum` VARCHAR(50) NULL,
  `SupplierName` VARCHAR(50) NULL,
  `Address` VARCHAR(50) NULL,
  `Phone` VARCHAR(50) NULL,
  `Fax` VARCHAR(50) NULL,
  `Notes` VARCHAR(50) NULL,
  PRIMARY KEY (`VenderID`),
  INDEX `VenderID` (`VenderID`)
)
ENGINE = INNODB;

DROP TABLE IF EXISTS `db2`.`Workorder Labor`;
CREATE TABLE `db2`.`Workorder Labor` (
  `WorkorderLaborID` INT(10) NOT NULL AUTO_INCREMENT,
  `WorkorderID` INT(10) NULL,
  `EmployeeID` INT(10) NULL,
  `BillableHours` DOUBLE(15, 5) NULL,
  `BillingRate` DECIMAL(19, 4) NULL,
  `Comment` VARCHAR(255) NULL,
  PRIMARY KEY (`WorkorderLaborID`),
  INDEX `{3BCB02AA-2E3D-448B-B3F9-8B0C2C` (`EmployeeID`),
  INDEX `{DF92C75C-67DA-49AB-9C13-1E80A0` (`WorkorderID`),
  INDEX `EmployeeID` (`EmployeeID`),
  INDEX `WorkorderID` (`WorkorderID`)
)
ENGINE = INNODB;

DROP TABLE IF EXISTS `db2`.`Workorder Parts`;
CREATE TABLE `db2`.`Workorder Parts` (
  `WorkorderPartID` INT(10) NOT NULL AUTO_INCREMENT,
  `WorkorderID` INT(10) NULL,
  `PartID` INT(10) NULL,
  `Quantity` INT(10) NULL,
  `UnitPrice` DECIMAL(19, 4) NULL,
  PRIMARY KEY (`WorkorderPartID`),
  INDEX `{5FC06EB0-388C-420D-A083-F5CFC9` (`WorkorderID`),
  INDEX `{C202B259-1A8B-4DC3-92B7-D987B7` (`PartID`),
  INDEX `PartID` (`PartID`),
  INDEX `WorkorderID` (`WorkorderID`),
  INDEX `WorkorderPartID` (`WorkorderPartID`)
)
ENGINE = INNODB;

DROP TABLE IF EXISTS `db2`.`Workorders`;
CREATE TABLE `db2`.`Workorders` (
  `WorkorderID` INT(10) NOT NULL AUTO_INCREMENT,
  `CustomerID` INT(10) NULL,
  `EmployeeID` INT(10) NULL,
  `PurchaseOrderNumber` VARCHAR(30) NULL,
  `DateReceived` DATETIME NULL,
  `DateRequired` DATETIME NULL,
  `MakeAndModel` VARCHAR(255) NULL,
  `SerialNumber` VARCHAR(50) NULL,
  `ProblemDescription` LONGTEXT NULL,
  `DateFinished` DATETIME NULL,
  `DatePickedUp` DATETIME NULL,
  `SalesTaxRate` DECIMAL(19, 4) NULL,
  `PaymentTerms` VARCHAR(50) NULL,
  `EstimateFooter` LONGTEXT NULL,
  `Status` VARCHAR(50) NULL,
  PRIMARY KEY (`WorkorderID`),
  INDEX `{6ED2FB1D-8F03-4011-925F-F9C2D4` (`CustomerID`),
  INDEX `CustomerID` (`CustomerID`),
  INDEX `DateFinished` (`DateFinished`),
  INDEX `DatePickedUp` (`DatePickedUp`),
  INDEX `EmployeeID` (`EmployeeID`),
  INDEX `SalesTaxRate` (`SalesTaxRate`),
  INDEX `SerialNumber` (`SerialNumber`)
)
ENGINE = INNODB;



SET FOREIGN_KEY_CHECKS = 1;

-- ----------------------------------------------------------------------
-- EOF

