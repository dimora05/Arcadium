@echo off
set CWD=%~dp0

set NewStr="C:\Thermo\SampleManager\Server\VGSM\Exe\"
set OldStr="SAMPLEMANAGERPATH"

set _FilePath=CustomizationObjectModel\
set _FileName=CustomizationObjectModel.csproj.config
set _OutFileName=CustomizationObjectModel.csproj.user

set _FilePath2=CustomizationTasks\
set _FileName2=CustomizationTasks.csproj.config
set _OutFileName2=CustomizationTasks.csproj.user

echo Configuring Customization Solution for Build
echo ============================================
echo Please set the path to the Instance Executable Directory

set /P NewStr="SampleManager Path [%NewStr%]:"

SETLOCAL
SETLOCAL ENABLEDELAYEDEXPANSION

if exist "%_FilePath%%_OutFileName%" (
    del "%_FilePath%%_OutFileName%" >nul 2>&1
    )

Call ConfigureSubstitute.bat %OldStr% %NewStr% %_FilePath%%_FileName% > %_FilePath%%_OutFileName%

if exist "%_FilePath2%%_OutFileName2%" (
    del "%_FilePath2%%_OutFileName2%" >nul 2>&1
    )

Call ConfigureSubstitute.bat %OldStr% %NewStr% %_FilePath2%%_FileName2% > %_FilePath2%%_OutFileName2%

:: Pause script for approx. 5 seconds...
echo Done
PING 127.0.0.1 -n 6 > NUL 2>&1
exit /b