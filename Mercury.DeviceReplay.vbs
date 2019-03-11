
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Class Name			:	DeviceReplay
'Description		:   Class to perform taks using the 'Mercury.DeviceReplay' object
'Assumptions		:   NA
'Functions			:   MouseMove(intX, intY)
'					:	MouseDblClick(intX, intY, intMouseButton)
'					:	MouseClick(intX, intY, intMouseButton)
'Author				:   Pratik R.
'Created Date		:   18-July-2016
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Class DeviceReplay

	Private objDevice
	Public LEFT_MOUSE_BUTTON, MIDDLE_MOUSE_BUTTON, RIGHT_MOUSE_BUTTON

	Private Sub Class_Initialize(  )
		Set objDevice = CreateObject ("Mercury.DeviceReplay")
		me.LEFT_MOUSE_BUTTON = LEFT_MOUSE_BUTTON
		me.MIDDLE_MOUSE_BUTTON = MIDDLE_MOUSE_BUTTON
		me.RIGHT_MOUSE_BUTTON = RIGHT_MOUSE_BUTTON
	End Sub

	Public Function MouseDown(byRef intX, byRef intY, byRef intMouseButton)
		objDevice.MouseDown intX, intY, intMouseButton
		oLogger.addLog "objDevice.MouseDown "&intX&", "&intY&", "&intMouseButton
	End Function

	Public Function MouseUp(byRef intX, byRef intY, byRef intMouseButton)
		objDevice.MouseUp intX, intY, intMouseButton
		oLogger.addLog "objDevice.MouseUp "&intX&", "&intY&", "&intMouseButton
	End Function

	Public Function MouseMove(byRef intX, byRef intY)
		objDevice.MouseMove intX, intY
		oLogger.addLog "objDevice.MouseMove "&intX&", "&intY
	End Function

	Public Function MouseDblClick(byRef intX, byRef intY, byRef intMouseButton)
		objDevice.MouseDblClick intX, intY, intMouseButton
		oLogger.addLog "objDevice.MouseDblClick "&intX&", "&intY&", "&intMouseButton
	End Function

	Public Function MouseClick(byRef intX, byRef intY, byRef intMouseButton)
		objDevice.MouseClick intX, intY, intMouseButton
		oLogger.addLog "objDevice.MouseClick "&intX&", "&intY&", "&intMouseButton
	End Function

	Public Function DragAndDrop(byRef intDragX, byRef intDragY, byRef intDropX, byRef intDropY, byRef intButton)
		objDevice.DragAndDrop intDragX, intDragY, intDropX, intDropY, intButton
		oLogger.addLog "objDevice.DragAndDrop "&intDragX&", "&intDragY&", "&intDropX&", "&intDropY
	End Function

	Private Sub Class_Terminate(  )
		Set objDevice = Nothing
	End Sub

End Class

Public oDevice
Set oDevice = new DeviceReplay
