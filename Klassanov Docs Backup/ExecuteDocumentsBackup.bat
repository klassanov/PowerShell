@echo off

for /f "delims=" %%a in ('wmic OS get localdatetime  ^| find "."') do set datetime=%%a


set "YYYY=%datetime:~0,4%"
set "MM=%datetime:~4,2%"
set "DD=%datetime:~6,2%"
set "HH=%datetime:~8,2%"
set "MI=%datetime:~10,2%"
set "SS=%datetime:~12,2%"

set fullstamp=%YYYY%-%MM%-%DD%@%HH%.%MI%.%SS%
Rem echo fullstamp=%fullstamp%

set  destDirName="D:\Backup %fullstamp%"
Rem echo destDirName=%destDirName%


mkdir %destDirName%
echo Directory %destDirName% created
pause


robocopy "C:\tmp\Source" %destDirName% /MIR
pause