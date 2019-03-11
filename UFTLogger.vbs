
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Class Name			:	Logger
'Description		:   Class to get print logs from script executions
'					:	saves logs to the local drive for debugging purposes
'Assumptions		:   NA
'Functions			:   addLog(strLogString) : 	To add print log to the UFT Print window
'					:							and also to the local file for debugging later
'Author				:   Pratik R.
'Created Date		:   23-June-2016
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Class Logger

	Private strLog
	Private strFileName
	Private objFSO

	Public Sub addLog(byRef strLogString)
		strLog = strLog & strLogString & vbNewLine
		Print strLogString
	End Sub

	Public Sub flushLogToFile()
		Set res = objFSO.OpenTextFile(strFileName, objFSO.ForAppending, True)
		res.Write strLog
		strLog = ""
		res.Close
		Set res = Nothing
	End Sub

	Private Sub Class_Initialize(  )
		strLog = ""
		Set objFSO = new FileSystem
		strLogFolder = Environment.Value("CommonFolderPath") & Environment.Value("Results")
		If not objFSO.FolderExists(strLogFolder) Then
			objFSO.CreateFolder strLogFolder
		End If
		arrTags = Split(Environment.Value("tag"), "\")
		For each strTag in arrTags
			strLogFolder = strLogFolder & "\" & strTag
			If not objFSO.FolderExists(strLogFolder) Then
				objFSO.CreateFolder strLogFolder
			End If
		Next
		strLogFolder = strLogFolder & "\" & Environment.Value("TestName")
		If not objFSO.FolderExists(strLogFolder) Then
			objFSO.CreateFolder strLogFolder
		End If
		strCurrentTime = Now()
		strCurrentTime = Replace(strCurrentTime, "/", "-")
		strCurrentTime = Replace(strCurrentTime, ":", "_")
		strLogFolder = strLogFolder & "\" & strCurrentTime
		If not objFSO.FolderExists(strLogFolder) Then
			objFSO.CreateFolder strLogFolder
		End If
		Environment.Value("strLogFolder") = strLogFolder
		strFileName = strLogFolder & "\PrintLog.txt"
		Set res = objFSO.CreateTextFile(strFileName, True)
		res.Write ""
		res.Close
		Set res = Nothing
	End Sub

	Private Sub Class_Terminate(  )
		flushLogToFile
		Set objFSO = Nothing
	End Sub

End Class

'Create logger object for use in the script
Public oLogger
Set oLogger = New Logger
