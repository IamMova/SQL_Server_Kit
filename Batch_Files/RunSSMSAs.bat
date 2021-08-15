:: Script Name		:   	RunSSMSAs.bat
:: Create Date		:	2021-08-15
:: Author		:   	Mova
:: Purpose		:   	To Run SQL Server Management Studio as Different user
:: Used By		:	MsSQL Developer
:: Source		:	https://github.com/IamMova/SQLServerPackage/tree/master/Batch_Files
:: -------------------------------------------------------------------------------------------------------
:: Revision History
:: -------------------------------------------------------------------------------------------------------
:: Date(yyyy-mm-dd)    Version/Ticket #		Author              	Comments
:: -------------------------------------------------------------------------------------------------------
:: 2021-08-15		1.0			Mova			Created
:: -------------------------------------------------------------------------------------------------------

@echo off

setlocal EnableDelayedExpansion
set programtorun=C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE\Ssms.exe
set _SSMS_USERKEY=

:: If needed, please update the SSMS path at line # 17

:main
	call :setconfiguration "%~0" %~1
	cls
	echo ==========================================
	echo   Run SQL Server Management Studio AS...
	echo ==========================================
	echo.
		
		if "!_SSMS_USERNAME!" == "" (
			set /p vUserName=_       Username : 
			set /p vPassword=_       Password : 
			set _SSMS_USERKEY=!vPassword!
		) else (
			set vUserName=!_SSMS_USERNAME!
			echo _       Username : !_SSMS_USERNAME!
			set /p vPassword=_       Password : 
			setx _SSMS_USERKEY=!vPassword!
		)
	echo.
	echo ==========================================
	
	if defined _ASK_ME_LATER ( 
		if "!_ASK_ME_LATER!" == "Y" ( call :saveusername %vUserName% )
	) else (
		call :saveusername %vUserName%
	)
	
	call :openappas %vUserName%
	
goto :eof

:openappas
	set iUserName=%~1
	set iPassword=!_SSMS_USERKEY!
	
	REM Prepare Temporary VB Script
	echo set WshShell = WScript.CreateObject(^"Wscript.Shell^") > runasprogram.vbs
	echo WshShell.run ^"runas /netonly /noprofile /user:!iUserName! ^" + Chr(34) + ^"!programtorun!^" + Chr(34) >> runasprogram.vbs
	echo WScript.Sleep 500 >> runasprogram.vbs
	echo WshShell.SendKeys ^"!iPassword!^" >> runasprogram.vbs
	echo WshShell.SendKeys ^"{ENTER}^" >> runasprogram.vbs
	echo set WshShell = nothing >> runasprogram.vbs
	
	REM Run prepared VB Script
	cscript /nologo runasprogram.vbs
	
	REM Delete VB SCript
	del /q runasprogram.vbs
	
	echo.
	cls

	if not "!errorlevel!" == "0" (
		call :showmessage "ERROR" "Failed" "Invalid username/Password..."
		goto :main
	) else (
		call :showmessage "SUCCESS" "Success" "Attemting to run SSMS as !iUserName!"..
		echo.
	)
goto :eof

:saveusername

	if not defined _SSMS_USERNAME (
		choice /m "Save username:"
		echo ==========================================
		if not "!errorlevel!" == "1" (
			choice /m "Ask me later:"
			
			if not "!errorlevel!" == "1" ( 
				setx _ASK_ME_LATER "N" > nul
				REM reg delete HKCU\Environment /F /V _ASK_ME_LATER
				REM setx _ASK_ME_LATER "" & reg delete HKCU\Environment /F /V _ASK_ME_LATER
			) else ( 
				setx _ASK_ME_LATER "Y" > nul
			)
			
		) else (
				setx _SSMS_USERNAME "%1" > nul
				setx _ASK_ME_LATER "N" > nul
		)
	)
goto :eof

:showmessage
	set messageType=%~1
	set messageTitle=%~2
	set messageText=%~3

	if /i "!messageType!" == "ERROR" (
		echo ER
		call :reconfiguration "FC"
	)
	if /i "!messageType!" == "SUCCESS" (
		echo SC
		call :reconfiguration "0A"
	)
	echo ==========================================
	echo                   !messageTitle!
	echo ==========================================
	echo.
	echo !messageText!
	echo.
	
	REM To pause the screen for 3 seconds
	timeout 3 > nul
	
goto :eof

:setconfiguration
	color B
	mode con: cols=43 lines=15
	
	if not exist "%cd%\ReadMe.txt" (
		echo To clear the stored username, You can type below command in cmd. >> ReadMe.txt
		echo. >> ReadMe.txt
		echo "%~1" /clear >> ReadMe.txt
		echo. >> ReadMe.txt
		echo It will clear the stored username and from next time it will start from fresh. >> ReadMe.txt
		echo. >> ReadMe.txt
		echo Thanks! >> ReadMe.txt
	)
	
	if "%~2" == "/clear" (
		reg delete HKCU\Environment /F /V _ASK_ME_LATER
		reg delete HKCU\Environment /F /V _SSMS_USERNAME
		
		exit > nul
	)
	
	if defined _SSMS_USERKEY ( reg delete HKCU\Environment /F /V _SSMS_USERKEY )
goto :eof

:reconfiguration
	color %~1
	mode con: cols=43 lines=10
goto :eof
