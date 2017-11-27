Attribute VB_Name = "Bas01ToDeveloper"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2014, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

'------------------------------------------------------------------------
'-- Components
'   * This one
'   * Rules.xlsx - includes information of first 3 sheets for all database types.
'   * TableData.xlsx - includes data for sample schema for all database types.
'   * Tests folder - includes test data
'------------------------------------------------------------------------

'------------------------------------------------------------------------
'-- How to build
'   * Run basBuild.Build
'-- Helper methods
'   * basBuildDataFileWorkbook.ExportToWorkbook
'       Create a new worksheet with data schema from the current workbook.
'------------------------------------------------------------------------

'------------------------------------------------------------------------
'-- Future Features
'   < Database Type >
'   App Gene
'
'   < Script Capability>
'   * Support user define type
'
'   < Other>
'   * More DDL features
'   * multiple lines for FK, Index
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
'   * Add a worksheet in Rules.xlsx
'   * Add a worksheet in TableData.xlsx
'   * Add a target test result on folder TestData
'   * Add constant for new databsae in basAppSetting File
'       E.g. Public Const DBName_Oracle                      As String = "Oracle"
'   * Add menu item for new database in basToolbar
'   * Add below code files
'       clsDB<NewDatabase>Provider
'       clsImportProvider<NewDatabase>
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

