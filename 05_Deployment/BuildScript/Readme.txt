The document is for deployment DB modeling excel.

Steps:
* Adjust source code excel's APP_VERSION information
* Adjust command in file 05_Deployment\Resources\Tools\Sample.bat
* Adjust build.vbs's version information.
* Except mdlExcelFunctions, clear all other macros and also code in the worksheet in template files.
* check build.vbs setting information, like filename, folder, vesion... etc.
* check 7-zip path in the build.vbs.
* run build.vbs
* get build result from .\00_ouput\deploy

*************************************************
Build Files list:
Build.vbs           : build main script
runExcelMacro.vbs   : functional script for invoke an excel's macro
*************************************************