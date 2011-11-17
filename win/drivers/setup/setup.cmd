::Switch to folder batch is executed from
cd /d %~dp0

::do setup stuff using curdir as path - everything is logically named
echo IP CONFIG>> %cd%\log.txt
cmd.exe /C cscript.exe %cd%\setstaticip.vbs >> %cd%\log.txt
echo METADATA>> %cd%\log.txt
cmd.exe /C cscript.exe %cd%\ec2userdata.vbs >> %cd%\log.txt

shutdown -r -f -t 5