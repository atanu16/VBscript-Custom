@echo off
echo Cleaning up temporary files...


:: Define log folder
set "logFolder=C:\TemClearLog"



:: Enable delayed expansion for proper variable evaluation inside the loop
setlocal enabledelayedexpansion

:: Check and clear old logs if the folder contains more than 30 files
if exist "%logFolder%" (
    set /a fileCount=0
    for %%F in ("%logFolder%\*.*") do (
        set /a fileCount+=1
    )
    if !fileCount! GTR 30 (
        echo Log folder contains more than 30 files. Deleting old logs...
        echo Log folder contains more than 30 files. Deleting old logs... >> "%logFolder%\cleanup_log.txt"
        del /q "%logFolder%\*.*" >nul 2>&1
        echo Old logs deleted. >> "%logFolder%\cleanup_log.txt"
        echo Old logs deleted.
    )
)

endlocal

:: Create log folder if it does not exist
if not exist "%logFolder%" (
    mkdir "%logFolder%"
)

:: Extract the current date in the desired format (DDMMYYYY)
for /f "tokens=1-4 delims=/.- " %%a in ('wmic os get localdatetime ^| find "."') do (
    set "currentDate=%%a"
)

:: Extract date and time components
set "formattedDate=%currentDate:~6,2%%currentDate:~4,2%%currentDate:~0,4%"
set "formattedTime=%currentDate:~8,2%%currentDate:~10,2%%currentDate:~12,2%"

:: Construct log file name
set "logFile=delete_log_%formattedDate%_%formattedTime%.txt"

:: Create log folder if it does not exist
if not exist "%logFolder%" (
    mkdir "%logFolder%"
)

:: Define the full path for the log file
set "logFilePath=%logFolder%\%logFile%"

:: Begin logging
echo Cleaning up temporary files... >> "%logFilePath%"
echo Start time: %date% %time% >> "%logFilePath%"
echo -------------------------------- >> "%logFilePath%"
echo Deleted files and folders: >> "%logFilePath%"

:: Delete temporary files in the system's temp directory
echo Scanning and deleting files in C:\Temp and %temp%...



for %%T in (%temp%;C:\Temp) do (
    echo ---------------------------------------------------------- >> "%logFilePath%"
    echo Scanning and deleting files in %%T... >> "%logFilePath%"
    
    :: Delete files in the directory
    for %%F in (%%T\*.*) do (
        echo Deleting file: %%F >> "%logFilePath%"
        del /q /f "%%F" >> "%logFilePath%" 2>&1
    )

    :: Delete folders in the directory
    for /d %%D in (%%T\*) do (
        echo Deleting folder: %%D >> "%logFilePath%"
        rmdir /s /q "%%D" >> "%logFilePath%" 2>&1
    )
)


:: Recreate Temp directories
echo -------------------------------- >> "%logFilePath%"
mkdir %temp% >> "%logFilePath%" 2>&1
mkdir C:\Temp >> "%logFilePath%" 2>&1

:: Log completion
echo Temporary files have been deleted.
echo -------------------------------- >> "%logFilePath%"
echo Temporary files have been deleted. >> "%logFilePath%"
echo Log saved to %logFilePath% >> "%logFilePath%"
echo Machine Restarted >> "%logFilePath%"


:: Pause before restarting
echo Restarting the computer in 5 seconds... 
timeout /t 6 >nul


:: Restart the computer
shutdown /r /t 0