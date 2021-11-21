@ECHO OFF
TITLE CLI: Microsoft Office - Online Installer v1.1.3
PUSHD "%~dp0"

(NET FILE||(powershell start-process -FilePath '%0' -verb runas)&&(exit /B)) >NUL 2>&1
(NET FILE||(exit)) >NUL 2>&1

ECHO                                      \\!//
ECHO                                      (o o)
ECHO ---------------------------------oOOo-(_)-oOOo---------------------------------
ECHO  This Microsoft Office 2019 Professional Plus installation includes:
ECHO  Word, Excel, PowerPoint, Outlook, OneNote, Publisher, Access, ProofingTools,
ECHO  Groove (OneDrive for Business), Lync (Skype for Business), Visio, Project.
ECHO ===============================================================================
ping -n 3 localhost 1>NUL


ECHO.
ECHO  After the Installation Run the 'KMS_VL_ALL - Smart Activation Script' from:
ECHO  http://www.zone62.com/downloads/programs/135-kms-vl-all-activator.
ECHO.
:INPUT-IYN
SET "OPTION="
SET /P "OPTION=-> Enter "1" to install Office 2019 Professional Plus or "0" to Quit: "
IF /I "%OPTION%"=="1" GOTO PAC
IF /I "%OPTION%"=="0" EXIT
GOTO INPUT-IYN

:PAC
IF NOT "%Processor_Architecture%" == "AMD64" GOTO x86

ECHO.
ECHO Installing Microsoft Office 2019 Professional Plus x64 (64-bit) . . .
START "" /WAIT /B ".\Microsoft Office Deployment Tool\setup.exe" /configure "Microsoft Office Deployment Tool\Office 2019 x64 Configuration.xml"
GOTO NEXT

:x86
ECHO.
ECHO Installing Microsoft Office 2019 Professional Plus x86 (32-bit) . . .
START "" /WAIT /B ".\Microsoft Office Deployment Tool\setup.exe" /configure "Microsoft Office Deployment Tool\Office 2019 x86 Configuration.xml"

:NEXT
ECHO.
ECHO Disabling Microsoft Office 2019 Telemetry . . .
REG ADD "HKLM\SOFTWARE\Microsoft\Office\Common\ClientTelemetry" /v "DisableTelemetry" /t REG_DWORD /d "00000001" /f 1>NUL


ECHO.
ECHO ----------------------------------[Finished]-----------------------------------
ping -n 5 localhost 1>NUL
EXIT
