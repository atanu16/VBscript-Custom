@echo off
echo Cleaning up temporary files...

:: Delete temporary files in the system's temp directory
echo Deleting files in %temp%...
del /q /f %temp%\*.*
rmdir /s /q %temp%

:: Delete temporary files in Windows Temp folder
echo Deleting files in C:\Windows\Temp...
del /q /f C:\Windows\Temp\*.*
rmdir /s /q C:\Windows\Temp

:: Recreate Temp directories
mkdir %temp%
mkdir C:\Windows\Temp

echo Temporary files have been deleted.
pause
