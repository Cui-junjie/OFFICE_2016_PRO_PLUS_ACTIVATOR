@echo off && setlocal enabledelayedexpansion
echo "It is Office 2016 Pro Plus Activator V1.0.0.1"
rem Initialization
rem 初始化变量
set officeFolder=""
rem Find office 2016
rem 寻找office2016
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" set officeFolder="%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" set officeFolder="%ProgramFiles(x86)%\Microsoft Office\Office16"
rem Find office 2016 fail
rem 找不到office2016
if %officeFolder%=="" call :End "Can't find Office 2016"
rem Find office 2016 successfully
rem 寻找office2016成功
echo "Find office 2016 successfully, Start detecting Office activation status"
rem Detect office activation status
rem 检测office激活状态
call :DetectOfficeStatus
rem Office has been activated
rem Office已被激活
echo "%statusInfo%"|findstr /i "licensed" 1>nul && echo "Office has been activated"
rem If Office is a retail version, Transfer Retail to Vol
rem Office为零售版，将office从零售版转为批量许可版
echo "%statusInfo%"|findstr /i "retail" 1>nul && call :RetailToVol
rem Activate office 2016
rem 激活office2016
call :Activation
goto End

:RetailToVol
rem Transfer from Retail to Vol
rem 将office从零售版转为批量许可版
echo "Transfer from Retail to Vol"
rem Install Vol version license
rem 安装office批量许可版许可证
set /a n=0
cscript  /nologo %officeFolder%\ospp.vbs /inslic:%officeFolder%\"..\root\Licenses16\ProPlusVL_KMS_Client-ppd.xrm-ms" | findstr /i "successfully" && set /a n+=1
cscript  /nologo %officeFolder%\ospp.vbs /inslic:%officeFolder%\"..\root\Licenses16\ProPlusVL_KMS_Client-ul.xrm-ms" | findstr /i "successfully" && set /a n+=1
cscript  /nologo %officeFolder%\ospp.vbs /inslic:%officeFolder%\"..\root\Licenses16\ProPlusVL_KMS_Client-ul-oob.xrm-ms" | findstr /i "successfully" && set /a n+=1
cscript  /nologo %officeFolder%\ospp.vbs /inslic:%officeFolder%\"..\root\Licenses16\ProPlusVL_MAK-pl.xrm-ms" | findstr /i "successfully" && set /a n+=1
cscript  /nologo %officeFolder%\ospp.vbs /inslic:%officeFolder%\"..\root\Licenses16\ProPlusVL_MAK-ppd.xrm-ms" | findstr /i "successfully" && set /a n+=1
cscript  /nologo %officeFolder%\ospp.vbs /inslic:%officeFolder%\"..\root\Licenses16\ProPlusVL_MAK-ul-oob.xrm-ms" | findstr /i "successfully" && set /a n+=1
cscript  /nologo %officeFolder%\ospp.vbs /inslic:%officeFolder%\"..\root\Licenses16\ProPlusVL_MAK-ul-phn.xrm-ms" | findstr /i "successfully" && set /a n+=1
if %n% lss 7  call :End "Fail to transfer"
rem Transfer office from Retail to Vol successfully
rem 成功将office从零售版转为批量许可版
echo "Transfer office from Retail to Vol successfully"
goto :eof

:Activation
rem Start to Activate office
rem 开始激活office
echo "Start to Activate office"
rem Install office product key
rem 安装office产品密钥
rem cscript  /nologo %officeFolder%\ospp.vbs /inpkey:JNRGM-WHDWX-FJJG3-K47QV-DRTFM
cscript  /nologo %officeFolder%\ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
rem Set the KMS server
rem 设置KMS服务器
rem cscript  /nologo %officeFolder%\ospp.vbs /sethst:kms.teevee.asia
cscript  /nologo %officeFolder%\ospp.vbs /sethst:kms.03k.org
rem rem Activate office
rem 激活office
cscript  /nologo %officeFolder%\ospp.vbs /act
rem Detect office activation status
rem 检测office激活状态
call :DetectOfficeStatus
echo "%statusInfo%"|findstr /i "licensed" 1>nul && call :End "Activate successfully"
rem Fail to activate
rem 激活失败
echo "Fail to activate"
goto :eof

:DetectOfficeStatus
rem Detect office activation status
rem 检测office激活状态
for /f "delims=" %%i in ('cscript /nologo %officeFolder%\ospp.vbs /dstatus') do (
	set statusInfo=!statusInfo!%%i
rem Show last 5 characters of installed product key
rem 显示已安装的产品密钥的后5位字符
	set currentLinestatusInfo=%%i
	if defined key (echo >nul) else (echo "!currentLinestatusInfo!" | findstr /i "key:")
)
goto :eof

:End
rem End program
rem 结束程序
set condition=%1
if defined condition echo !condition!
pause
exit

:Clearkey
rem Delete useless key
rem 删除无用密钥
echo "Delete useless key"
if "%key%" neq "" cscript  /nologo %officeFolder%\ospp.vbs /unpkey:%key%
goto :eof

rem Apply to Office 2016 Pro Plus on Windows