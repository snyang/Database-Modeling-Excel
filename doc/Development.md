# Development Information

## Release file list:

[release]

```txt
DB_Modeling_Excel_<version>.zip
DB_Modeling_Excel\doc\Help.docx
DB_Modeling_Excel\tools\DBME_CopyToFile.vbs
DB_Modeling_Excel\tools\DBME_RunExcelMacro.vbs
DB_Modeling_Excel\tools\Sample.bat
```

## Develop with github

### Source code

**For most cases, we only need to change files which are under \src\.**

- \src\DME_Template_All_7_0.xlsm

Almost all source code are in the file.

- \src\DB_TableDefinitions.xlsx

Include all table definition data for generated DME_Template_*_<version>.xlsm.

The build process will import these data into these Excel files.

- \src\DB_Rules.xlsx

Include content for the sheet "~Rules~" of all generated DME_Template_*_<version>.xlsm.

The build process will import these data into these Excel files.

- \src\doc\Help.docx

The user guide.

- \src\tools\DBME_CopyToFile.vbs
- \src\tools\DBME_RunExcelMacro.vbs
- \src\tools\Sample.bat

- src\macro

DO NOT change these files manually. These files are generated during builing, and just provide a way to know what we are changing in the spreadsheet.

Small utilities for their DevOps processes.

Includes commands to generate SQL script files from the database model excel.

### Check out
Please use the branch develop or a private branch to your developing

```bat
git checkout -b develop master
```

### Build and Publish

- Finish your your changes and testing
- Change product version if need
- Enable "Trust access the VBA project object model"  
  Excel > File > Excel Options > Trust Center  
  Trust Center > Macro Settings  
  Check "Trust access the VBA project object model"  
- Build and Test
  Run DME_Template_*_<version>.xlsm's macro: basBuild.Build.
- git commit and publish

```bat
git push -u origin develop
```

## Publish to sourceforge.net

Login yang_ning  
goto the project: https://sourceforge.net/projects/db-model-excel  
Select menu Files  
Go to folder "database modeling excel"  
Add a new folder, e.g. 5.0.0 Production Release  
Go to the new folder  
Upload files into the folder. Done.  

Update project summary, changed version information.
