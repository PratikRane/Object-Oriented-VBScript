'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Class Name			:	FileSystem
'Description		:   Class to handle FSO tasks
'Assumptions		:   NA
'Functions			:   FileExists(strFilePath)
'					:	FolderExists(strFolderPath)
'					:	CreateFolder(strFolderPath)
'					:	GetFolder(strFolderPath)
'					:	DeleteFolder(strFolderPath)
'					:	CreateTextFile(strTXTfilePath, boolState)
'					:	DeleteFile(strFilePath,  boolForce)
'					:	GetFile(strFilePath)
'					:	CopyFile(strSourceFilePath, strDestinationFilePath, boolOverwrite)
'Author				:   Pratik R.
'Created Date		:   18-July-2016
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Class FileSystem

	Private objFSO
	Public ForReading, ForWriting, ForAppending	'IO Mode Constants
	Public TristateTrue, TristateFalse, TristateUseDefault	'Tristate Constants

	Private Sub Class_Initialize(  )
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		me.ForReading = 1
		me.ForWriting = 2
		me.ForAppending = 8
		me.TristateTrue = -1
		me.TristateFalse = 0
		me.TristateUseDefault = -2
	End Sub

	Public Function FileExists(byRef strFilePath)
		FileExists = objFSO.FileExists(strFilePath)
	End Function

	Public Function FolderExists(byRef strFolderPath)
		FolderExists = objFSO.FolderExists(strFolderPath)
	End Function

	Public Function CreateFolder(byRef strFolderPath)
		CreateFolder = objFSO.CreateFolder(strFolderPath)
	End Function

	Public Function GetFolder(byRef strFolderPath)
		Set GetFolder = objFSO.GetFolder(strFolderPath)
	End Function

	Public Function DeleteFolder(byRef strFolderPath)
		objFSO.DeleteFolder strFolderPath
	End Function

	Public Function CreateTextFile(byRef strTXTfilePath, byRef boolState)
		Set CreateTextFile = objFSO.CreateTextFile(strTXTfilePath, boolState)
	End Function

	Public Function OpenTextFile(byRef strTXTfilePath, byRef intState, byRef boolState)
		Set OpenTextFile = objFSO.OpenTextFile(strTXTfilePath, intState, boolState)
	End Function

	Public Function DeleteFile(byRef strFilePath, byRef boolForce)
		objFSO.DeleteFile strFilePath, boolForce
	End Function

	Public Function GetFile(byRef strFilePath)
		Set GetFile = objFSO.GetFile(strFilePath)
	End Function

	Public Function CopyFile(byRef strSourceFilePath, byRef strDestinationFilePath, byRef boolOverwrite)
		objFSO.CopyFile strSourceFilePath, strDestinationFilePath, boolOverwrite
	End Function

	Private Sub Class_Terminate(  )
		Set objFSO = Nothing
	End Sub

End Class

Public oFSO
Set oFSO = new FileSystem

