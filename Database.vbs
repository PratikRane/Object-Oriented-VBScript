
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Class Name			:	Database
'Description		:   Class to get print logs from script executions
'					:	saves logs to the local drive for debugging purposes
'Assumptions		:   NA
'Functions			:   runQuery(strQuery) 	: 	To run query and populate the resultset
'					:	Field(strField)		:	To retreive the value for the DB field
'					:	nextRow				:	To move to the next row in the result
'					:	EOF					:	To check if the end of DB rows is reached
'Author				:   Pratik R.
'Created Date		:   23-June-2016
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Class Database

	'Set up the properties needed for DB connection
	Private objConnection, ResultSet
	Public connected, queryRun
	Public strConnectString

	'Setup the connection and result set object for use before querying
	Private Sub Class_Initialize(  )
		connected = False
		queryRun = False
		Set objConnection = CreateObject("adodb.connection")
		Set ResultSet = CreateObject("adodb.recordset")
		strConnectString = ""
	End Sub

	'Connect to the DB using the provided connection string
	Public Sub connectDB(byRef strConnectionString)
		'End previous connection if any
		If connected Then
			If ResultSet.State = 1 Then ResultSet.Close
			If objConnection.State = 1 Then objConnection.Close
			oLogger.addLog "Closing previous DB Connection"
		End If
		'Try connecting to the database
		On Error Resume Next
		Err.Clear
		objConnection.Open strConnectionString
		On Error Goto 0
		'Check if DB connection was made
		If objConnection.State <> 1 Then
			oLogger.addLog "DB Connection failed. Check if VPN is connected"
			Exit Sub
		Else
			oLogger.addLog "DB connection successfull"
			connected = True
		End If
	End Sub

	'Execute the provided SQL query and return True/False if the query returned any results
	Public Function runQuery(byRef strQuery)
		If Not connected Then	'If DB is not connected, try connecting DB
			'Connect only if connection string is provided
			If strConnectString = "" Then
				oLogger.addLog "Query ResultSet cannot be created. No Connection string provided."
				Reporter.ReportEvent micFail, "Database query not completed", "Query ResultSet cannot be created. No Connection string provided."
				ExitRun
			Else
				connectDB(strConnectString)
			End If
		End If
		If Not connected Then	'If DB is still not connected, Fail and exit
			oLogger.addLog "Query ResultSet cannot be created. No Database connected."
			Reporter.ReportEvent micFail, "Database query not completed", "Query ResultSet cannot be created. No Database connected."
			ExitRun
		End If
		If queryRun Then		'If query was run previously, close the result set for reusing the object
			ResultSet.Close
		End If
		'Run the SQL Query
		oLogger.addLog "Running following query:"&chr(10)&strQuery&chr(10)&"Query running..."&chr(10)&"Please wait..."
		ResultSet.Open strQuery, objConnection
		queryRun = True
		If ResultSet.EOF Then
			runQuery = False
		Else
			runQuery = True
		End If
		oLogger.addLog "DB query completed running. ResultSet created"
	End Function

	'Return the value of the field in the current DB result row
	Public Default Property Get Field(byRef strField)
		Field = ResultSet.Fields(strField).Value
		oLogger.addLog "oDB.ResultSet.Field('"&strField&"') = "& Field
	End Property

	'Move to the next row in the DB result set
	'If on the last row, stay on the last row
	Public Sub nextRow()
		If Not ResultSet.EOF Then ResultSet.MoveNext	
	End Sub

	'Check if DB result set is pointing to the last row
	Public Function EOF()
		EOF = ResultSet.EOF
	End Function

	'Close the DB connection and result set if open at the end of the script
	Private Sub Class_Terminate(  )
		If connected Then
			If ResultSet.State = 1 Then ResultSet.Close
			If objConnection.State = 1 Then objConnection.Close
			oLogger.addLog "Closing DB Connection"
		End If
		Set ResultSet = Nothing
		Set objConnection = Nothing
	End Sub

End Class

'Create Database objects for use in the script
Public oDB, oDB2
Set oDB = New Database
oDB.strConnectString = "Driver={Microsoft ODBC for Oracle};Server="&Environment.Value("DBName")&"; Uid="&Environment.Value("DBUsername")&";Pwd="&Environment.Value("DBPassword")&";"
Set oDB2 = New Database
