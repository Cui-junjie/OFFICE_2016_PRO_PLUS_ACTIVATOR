@echo off && setlocal enabledelayedexpansion
echo "Office 2016 Pro Plus Activator V1.0.0.0 For first time activation"
::Initialization
::初始化变量
set officeFolder=""
::Find office 2016
::寻找office2016
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" set officeFolder="%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" set officeFolder="%ProgramFiles(x86)%\Microsoft Office\Office16"
::Find office 2016 fail
::找不到office2016
if %officeFolder%=="" call :End "Can't find Office 2016"
::Find office 2016 successfully
::寻找office2016成功
echo "Find office 2016 successfully, Start detecting Office activation status"
::Detect office activation status
::检测office激活状态
call :DetectOfficeStatus
::Office has been activated
::Office已被激活
echo "%statusInfo%"|findstr /i "licensed" 1>nul && call :End "Office has been activated"
::If Office is a retail version, Transfer Retail to Vol
::Office为零售版，将office从零售版转为批量许可版
echo "%statusInfo%"|findstr /i "retail" 1>nul && call :RetailToVol
::Activate office 2016
::激活office2016
call :Activation
goto End

:RetailToVol
::Transfer from Retail to Vol
::将office从零售版转为批量许可版
echo "Transfer from Retail to Vol"
::Install Vol version license
::安装office批量许可版许可证
set /a n=0
cscript  /nologo %officeFolder%\ospp.vbs /inslic:%officeFolder%\"..\root\Licenses16\ProPlusVL_KMS_Client-ppd.xrm-ms" | findstr /i "successfully" && set /a n+=1
cscript  /nologo %officeFolder%\ospp.vbs /inslic:%officeFolder%\"..\root\Licenses16\ProPlusVL_KMS_Client-ul.xrm-ms" | findstr /i "successfully" && set /a n+=1
cscript  /nologo %officeFolder%\ospp.vbs /inslic:%officeFolder%\"..\root\Licenses16\ProPlusVL_KMS_Client-ul-oob.xrm-ms" | findstr /i "successfully" && set /a n+=1
cscript  /nologo %officeFolder%\ospp.vbs /inslic:%officeFolder%\"..\root\Licenses16\ProPlusVL_MAK-pl.xrm-ms" | findstr /i "successfully" && set /a n+=1
cscript  /nologo %officeFolder%\ospp.vbs /inslic:%officeFolder%\"..\root\Licenses16\ProPlusVL_MAK-ppd.xrm-ms" | findstr /i "successfully" && set /a n+=1
cscript  /nologo %officeFolder%\ospp.vbs /inslic:%officeFolder%\"..\root\Licenses16\ProPlusVL_MAK-ul-oob.xrm-ms" | findstr /i "successfully" && set /a n+=1
cscript  /nologo %officeFolder%\ospp.vbs /inslic:%officeFolder%\"..\root\Licenses16\ProPlusVL_MAK-ul-phn.xrm-ms" | findstr /i "successfully" && set /a n+=1
if %n% lss 7  call :End "Fail to transfer"
::Transfer office from Retail to Vol successfully
::成功将office从零售版转为批量许可版
echo "Transfer office from Retail to Vol successfully"
goto :eof

:Activation
::Start to Activate office
::开始激活office
echo "Start to Activate office"
::Install office product key
::安装office产品密钥
::cscript  /nologo %officeFolder%\ospp.vbs /inpkey:JNRGM-WHDWX-FJJG3-K47QV-DRTFM
cscript  /nologo %officeFolder%\ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
::Set the KMS server
::设置KMS服务器
::cscript  /nologo %officeFolder%\ospp.vbs /sethst:kms.teevee.asia
cscript  /nologo %officeFolder%\ospp.vbs /sethst:kms.03k.org
::::Activate office
::激活office
cscript  /nologo %officeFolder%\ospp.vbs /act
::Detect office activation status
::检测office激活状态
call :DetectOfficeStatus
echo "%statusInfo%"|findstr /i "licensed" 1>nul && call :End "Activate successfully"
::Fail to activate
::激活失败
echo "Fail to activate"
goto :eof

:DetectOfficeStatus
::Detect office activation status
::检测office激活状态
for /f "delims=" %%i in ('cscript  /nologo %officeFolder%\ospp.vbs /dstatus') do (
	set statusInfo=!statusInfo!%%i
::	if "!currentLinestatusInfo!" equ "<No installed product keys detected>" (
::		set currentLinestatusInfo=retail
::		set statusInfo=!statusInfo!!currentLinestatusInfo!
::	)
::Show last 5 characters of installed product key
::显示已安装的产品密钥的后5位字符
	set currentLinestatusInfo=%%i
	if defined key (echo >nul) else (echo "!currentLinestatusInfo!" | findstr /i "key:")
::	if defined key (echo >nul) else (echo "!currentLinestatusInfo!" | findstr /i "key:" && for /f "tokens=2 delims=:" %%j in ("!currentLinestatusInfo!") do (
::		set key=%%j
::		set key=!key:~1!
::		)
::	)
)
goto :eof

:End
::End program
::结束程序
set condition=%1
if defined condition echo !condition!
pause
exit

:Clearkey
::Delete useless key
::删除无用密钥
echo "Delete useless key"
if "%key%" neq "" cscript  /nologo %officeFolder%\ospp.vbs /unpkey:%key%
goto :eof

::Apply to Office 2016 Pro Plus on Windows