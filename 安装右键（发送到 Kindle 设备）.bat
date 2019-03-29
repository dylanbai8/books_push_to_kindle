@echo off
cd /d "%~dp0"

::命令行标题栏和文字颜色
title -- 安装右键（发送到 Kindle 设备） --
MODE con: COLS=46 lines=12
color 0a

goto menu

:menu
MODE con: COLS=46 lines=12
set id=
set kindle_email=
set send_email=
set send_email_password=
cls
echo.
echo. ===== 安装右键（发送到 Kindle 设备）=====
echo.
echo.    1.安装到右键“发送到”菜单
echo.
echo.    2.从右键“发送到”菜单卸载
echo.
echo.    使用方法：选中电子书-右键-发送到
echo.
set /p id=请输入选择的序号:
if "%id%"=="1" goto install
if "%id%"=="2" goto unstall else (
goto menu
)

:install
MODE con: COLS=46 lines=24
cls
echo.
echo. ===== 安装右键（发送到 Kindle 设备）=====
echo.

::设置邮箱信息
echo. 1.设置 Kindle 邮箱（接收邮箱）
echo.
set /p kindle_email=请输入:
if not defined kindle_email goto error
echo.

echo. 2.设置用于发送附件的邮箱
echo.
set /p send_email=请输入:
if not defined send_email goto error
echo.

for /f "tokens=2,3,4 delims=@." %%a in ("%send_email%") do (
set h1=%%a
set h2=%%b
set h3=%%c
)

if not defined h3 (set domain=%h1%.%h2%) else (set domain=%h2%.%h3%)

echo. 3.设置该邮箱的SMTP服务器
echo. （默认服务器：smtp.%domain%）
echo.
set /p send_email_smtp=请输入:
if not defined send_email_smtp (set send_email_smtp=smtp.%domain%)
echo.

echo. 4.设置该邮箱的授权码（或密码）
echo. （将被明文储存在 send-to-kindle.vbs 中）
echo.
set /p send_email_password=请输入:
if not defined send_email_password goto error
echo.

::生成vbs程序
del /f /q %AppData%\Microsoft\Windows\SendTo\发送到-Kindle-设备.lnk 1>nul 2>nul
rd /s /q %cd%\app 1>nul 2>nul
md %cd%\app

(echo dim strFilepath
echo strFilepath = WScript.Arguments(0^)
echo kname = "《" ^& mid(strFilepath,instrrev(strFilepath,"\"^)+1^) ^& "》"
echo Currentdate = date(^)
echo ktime = "推送时间：" ^& Currentdate
echo.
echo.
echo '设置 收件邮箱
echo Const Email_To = "%kindle_email%"
echo.
echo '设置 发件邮箱
echo Const Email_From = "%send_email%"
echo Const Password = "%send_email_password%"
echo.
echo.
echo Set CDO = CreateObject("CDO.Message"^)
echo CDO.Subject = "电子书:" ^& kname ^& ktime
echo CDO.From = Email_From
echo CDO.To = Email_To
echo CDO.TextBody = "提示：请查阅附件。"
echo CDO.AddAttachment strFilepath
echo Const schema = "http://schemas.microsoft.com/cdo/configuration/"
echo With CDO.Configuration.Fields
echo 	.Item(schema ^& "sendusing"^) = 2
echo 	.Item(schema ^& "smtpserver"^) = "%send_email_smtp%"
echo 	.Item(schema ^& "smtpauthenticate"^) = 1
echo 	.Item(schema ^& "sendusername"^) = Email_From
echo 	.Item(schema ^& "sendpassword"^) = Password
echo 	.Item(schema ^& "smtpserverport"^) = 465
echo 	.Item(schema ^& "smtpusessl"^) = True
echo 	.Item(schema ^& "smtpconnectiontimeout"^) = 60
echo 	.Update
echo End With
echo CDO.Send
echo.
echo MsgBox ktime ^& Chr(13^) ^& "推送到：" ^& Email_To ^& Chr(13^) ^& "推送文件:" ^& kname, vbOKOnly, "推送成功")>%cd%\app\send-to-kindle.vbs

::生成快捷方式并安装到“发送到”目录
set Program=%cd%\app\send-to-kindle.vbs
set LnkName=发送到-Kindle-设备
set WorkDir=%cd%\app
set Desc=用于右键“发送到”菜单

(echo Set WshShell=CreateObject("WScript.Shell"^)
echo Set oShellLink=WshShell.CreateShortcut("%cd%\app\%LnkName%.lnk"^)
echo oShellLink.TargetPath="%Program%"
echo oShellLink.WorkingDirectory="%WorkDir%"
echo oShellLink.WindowStyle=1
echo oShellLink.Description="%Desc%"
echo oShellLink.IconLocation="%SystemRoot%\System32\SHELL32.dll,192"
echo oShellLink.Save)>%cd%\app\makelnk.vbs

%cd%\app\makelnk.vbs
del /f /q %cd%\app\makelnk.vbs 1>nul 2>nul
copy %cd%\app\%LnkName%.lnk %AppData%\Microsoft\Windows\SendTo 1>nul 2>nul

MODE con: COLS=46 lines=8
cls
echo.
echo. ===== 安装右键（发送到 Kindle 设备）=====
echo.
echo. 安装中，请稍后。。。
echo.
echo. 安装完成！按任意键退出安装程序
pause 1>nul 2>nul
exit

:unstall
MODE con: COLS=46 lines=8
cls
echo.
echo. ===== 安装右键（发送到 Kindle 设备）=====
echo.
echo. 卸载中，请稍后。。。
del /f /q %AppData%\Microsoft\Windows\SendTo\发送到-Kindle-设备.lnk 1>nul 2>nul
rd /s /q %cd%\app 1>nul 2>nul
echo.
echo. 卸载完成！按任意键退出安装程序
pause 1>nul 2>nul
exit

:error
MODE con: COLS=46 lines=8
cls
echo.
echo. ===== 安装右键（发送到 Kindle 设备）=====
echo.
echo. 错误！该值不能为空！
echo.
echo. 将在 3 秒后自动返回开始菜单！
ping localhost -n 3 1>nul 2>nul
goto menu