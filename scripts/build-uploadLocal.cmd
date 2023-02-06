
rem all paths are relative to the git scripts folder

set appName=menucrm0210

call build.cmd

rem upload to contensive application
c:
cd %collectionPath%
cc -a %appName% --installFile "%collectionName%.zip"
cd ..\..\scripts

pause