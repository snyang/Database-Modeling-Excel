Attribute VB_Name = "bas00Document"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2014, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

'------------------------------------------------------------------------
'-- ! Features
'-- * Supported Databases
'-- ** DB2
'-- ** MariaDB
'-- ** MySQL
'-- ** Oracle
'-- ** PostgreSQL
'-- ** SQL Server
'-- ** SQLite (Only support generating DDL feature)
'-- * Design and maintenance database schema
'-- * Generate script of database schema using workbook’s content.
'-- ** Create Tables SQL
'-- ** Drop Tables SQL
'-- ** Drop and Create Tables SQL
'-- ** Create Tables IF Not Exists SQL
'-- * Support ignore some worksheets when generating SQL scripts
'-- * Import database schema from database into the workbook
'-- * Support automatic build process
'-- * Fine print (A4 paper)
'--
'-- ! Enhancement and defect fixed list
'-- !!  <7.0>
'-- * New Features
'-- **  New UI style.
'-- **  Support DB2.
'-- **  Support MariaDB.
'-- *  Fixed Defects
'-- **  Fixed some defects.
'-- **  Refined generated SQL.
'-- **  Refined some code.
'--
'-- !!  <5.0.1>
'-- * New Features
'-- **  Support import Oracle comments from tables and columns.
'-- **  Support add hyperlinks to the table sheets in the index sheet.
'-- *  Fixed Defects
'-- **  Oracle: Error when connect Oracle during importing.
'-- **  Import: In import dialog, it should display 'Next >' when selecting tables.
'--
'-- !!  <5.0.0>
'-- * New Features
'-- **  Support PostgreSQL
'-- *  Fixed Defects
'-- **  During importing, the last row disappears in some cases.
'-- **  SQLite: Not "IF NOT EXISTS" for CREATE INDEX as creating IF NOT EXISTS SQL.
'--
'-- !!  <4.0.0>
'-- * New Features
'-- ** Support SQLite
'-- ** Oracle: Enhance the Import UI.
'
'-  <3.2.3>
'*  Bugs fix:
'*      [SQL Server] Cannot get FK when the reversing table is in a schema rather than dbo.
'
'-  <3.2.2>
'*  Bugs fix:
'*      A mistake in sample.bat
'
'-  <3.2.0>
'*  New feature:.
'*      Support only Create Table SQL script
'
'-  <3.1.0>
'*  New feature:.
'*      Support ignore some worksheets when generating SQL scripts
'
'-  <3.0.0 RC1>
'*  Support Oracle.
'*  Add a vb scripts sample which generate scripts
'
'-  <2.0.0 RC1>
'*  Support MySQL.
'*  Add a vb scripts sample which generate scripts
'*  More details Help
'*  Add denote function
'*  Per document for one database type
'
'-  <Bug Fixed>
'*  Generate Foreign Key Error.
'*  Generate Index bug: non-unique index is as unique index.
'*  Reverse Index bug: error when judgment an index is unique or not.
'*  Reverse index bug: error when there are more than one column in an index.
'
'-  <1.6.2006.0814>
'   Clear note when reverse
'   Default (1) change to -1 when reverse
'------------------------------------------------------------------------

'------------------------------------------------------------------------
'-- Current Features
'   < Database Type >
'   * Support SQL Server
'   * Support MySQL
'   * Oracle
'   * SQLite (Only has generating DDL feature)
'------------------------------------------------------------------------

'------------------------------------------------------------------------
'-- Future Features
'   < Database Type >
'   PostgreSQL
'   DB2
'
'   < Script Capability>
'   * Support user define type
'
'   < Other>
'   * More DDL features
'   * Unify importing code
'   * Generate unified XML
'   * multiple lines for FK, Index
'   * Better UI styles/themes for spreadsheet
'------------------------------------------------------------------------

'------------------------------------------------------------------------
'-- How to define version
'   * MajorVersion is changed for big feature be added.
'   * MinorVersion is changed for normal feature be added.
'   * Revision number is changed for bug fix.
'   * Release type definition. a#: arlfa release; b#: beta release; rc#:realease cadidate; <empty>: product release
'------------------------------------------------------------------------

'------------------------------------------------------------------------
'-- How to support a new database
'   * Add constant for new databsae in basAppSetting File
'       E.g. Public Const DBName_Oracle                      As String = "Oracle"
'   * Add menu item for new database in basToolbar
'   * Add below code files
'       clsDB<NewDatabase>Provider
'       basImport_<NewDatabase>
'       frmImport_<NewDatabase>
'   * Add select case for the database in basPublicDatabase.GetDatabaseProvider
'   * Update code in basUpgrade.SetTheExcelTypeVariable
'   * Add a template excel file in 05_Deployment\Resources
'       DatabaseModeling_Template_<NewDatabase>.xls
'   * Update build script for update macro for the new template excel file.
'
'   [Technical points]
'       Create Table
'       Create PK(Cluster, Unique, Non-unique), Index(Cluster, Unique, Non-unique), FK,
'       Create Table and Field comment
'       Drop Table
'       Drop FK
'       Get Table(s) schema information
'------------------------------------------------------------------------

'------------------------------------------------------------------------
'-- UI Guidelines
'-- * 8px between a container with the controls inside
'--   A container is one of Form, MultiPage
'-- * 4px between controls in horizontal/verticle direction
'-- * Generatl
'--   Height is 18px.
'-- * Button
'--   Height: 20px, Width: 72 px at least

'------------------------------------------------------------------------
