
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Class Name			:	wScriptShell
'Description		:   Class to perform taks on the 'WScript.Shell' object
'Assumptions		:   NA
'Functions			:   SendKeys(strKeyStrokes)								: 	To send keyboard entries to the application
'					:	Exec(strExecCommand)								:	To execute cmd commands line using EXEC function
'					:	Run(strRunCommand, intWindowStyle, boolWaitOnReturn):	To run cmd commands line using RUN function
'					:	ExpandEnvironmentStrings(strEnvString)				:	To return expanded strings for system environment variables
'Author				:   Pratik R.
'Created Date		:   18-July-2016
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Class wScriptShell

	Private oShell
	Public REG_SZ, REG_EXPAND_SZ, REG_DWORD, REG_BINARY
	Public HKEY_CURRENT_USER, HKCU, HKEY_USERS, HKEY_LOCAL_MACHINE, HKLM, HKEY_CLASSES_ROOT, HKCR, HKEY_CURRENT_CONFIG

	Private Sub Class_Initialize(  )
		Set oShell = CreateObject("WScript.Shell")
		REG_SZ = "REG_SZ"
		REG_EXPAND_SZ = "REG_EXPAND_SZ"
		REG_DWORD = "REG_DWORD"
		REG_BINARY = "REG_BINARY"
		HKEY_CURRENT_USER = "HKEY_CURRENT_USER"
		HKCU = "HKCU"
		HKEY_USERS = "HKEY_USERS"
		HKEY_LOCAL_MACHINE = "HKEY_LOCAL_MACHINE"
		HKLM = "HKLM"
		HKEY_CLASSES_ROOT = "HKEY_CLASSES_ROOT"
		HKCR = "HKCR"
		HKEY_CURRENT_CONFIG = "HKEY_CURRENT_CONFIG"
	End Sub

	Public Function SendKeys(byRef strKeyStrokes)
		oShell.SendKeys strKeyStrokes
	End Function

	Public Function Exec(byRef strExecCommand)
		oShell.Exec strExecCommand
	End Function

	Public Function Run(byRef strRunCommand, byRef intWindowStyle, byRef boolWaitOnReturn)
		Run = oShell.Run(strRunCommand, intWindowStyle, boolWaitOnReturn)
	End Function

	Public Function ExpandEnvironmentStrings(byRef strEnvString)
		ExpandEnvironmentStrings = oShell.ExpandEnvironmentStrings(strEnvString)
	End Function

	Public Function RegWrite(byRef strRegName, byRef anyValue, byRef strType)
		oShell.RegWrite strRegName, anyValue, strType
	End Function

	Public Function RegRead(byRef strRegName)
		RegRead = oShell.RegRead(strRegName)
	End Function

	Public Function RegDelete(byRef strRegName)
		oShell.RegDelete(strRegName)
	End Function

	Public Function AppActivate(byRef app)
		oShell.AppActivate app
	End Function

	Public Function CurrentDirectory()
		CurrentDirectory = oShell.CurrentDirectory()
	End Function

	Public Function LogEvent(byRef oType, byRef strMessage, byRef oTarget)
		oShell.LogEvent oType, strMessage, oTarget
	End Function

	Public Function Popup(byRef strTextMsg, byRef intSecondsToWait, byRef strTitle, byRef intType)
		oShell.Popup strTextMsg, intSecondsToWait, strTitle, intType
	End Function

	Private Sub Class_Terminate(  )
		Set oShell = Nothing
	End Sub

End Class

Public wShell
Set wShell = new wScriptShell
