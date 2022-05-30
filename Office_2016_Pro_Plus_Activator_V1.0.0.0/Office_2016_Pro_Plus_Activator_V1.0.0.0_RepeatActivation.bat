@echo off && setlocal enabledelayedexpansion
echo "It is Office 2016 Pro Plus Activator V1.0.0.0 for repeat activation"
::Go to the Office2016 installation path
::转到office安装路径
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd "%ProgramFiles(x86)%\Microsoft Office\Office16"

::Install the secret key
::安装产品密钥
::cscript  /nologo %officeFolder%\ospp.vbs /inpkey:JNRGM-WHDWX-FJJG3-K47QV-DRTFM
cscript /nologo ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99

::Set the KMS server
::设置KMS服务器
::cscript  /nologo %officeFolder%\ospp.vbs /sethst:kms.teevee.asia
cscript /nologo ospp.vbs /sethst:kms.03k.org

::Activate
::激活office
cscript /nologo ospp.vbs /act

::Detect the Activation Status
::检测激活状态
::cscript /nologo ospp.vbs /dstatus
for /f "delims=" %%i in ('cscript  /nologo ospp.vbs /dstatus') do ( set statusInfo=!statusInfo!%%i )
echo "%statusInfo%"|findstr /i "licensed" 1>nul && call :End "Activate successfully"

:End
::End program
::结束程序
set condition=%1
if defined condition echo !condition!
pause
exit

::Apply to Office 2016 Pro Plus on Windows