'
' Each module needs to have:
' Class_Initialize, define the Name
' Work and List_Options Sub
'
'
Register_Class( New LogicalDisk)


Class LogicalDisk
	' Define Properties
	Public Name
	Public Options
	Private WMI_Name
	
	' Initialize the Class
	Private Sub Class_Initialize
		Name = "LogicalDisk"
		WMI_Name = "Win32_LogicalDisk"
	End Sub
	
	Public Function Work (ComputerName, SelectedOptionsArray)
		'MsgBox "[" & ComputerName & "]"
		Dim WMIService, CollectionItems, ResultArray
		ResultArray = array()
		
			Set WMIService = GetObject("winmgmts:\\" & ComputerName & "\root\cimv2")
			Set CollectionItems = WMIService.ExecQuery("Select * from " & WMI_Name)
			For Each objItem in CollectionItems
				
				For Each SelectedOption in SelectedOptionsArray
					'MsgBox objItem.Properties_(SelectedOption).Value & " " & TypeName(objItem.Properties_(SelectedOption).Value)
					'MsgBox objItem.Properties_("asdas")
					'MsgBox "test" & objItem.Properties_(SelectedOption)
					
					CurrentProperty = objItem.Properties_(SelectedOption).Value
					'On Error Resume Next
					If (Err.Number <> 0) Then
						ReDim Preserve ResultArray(UBound(ResultArray)+1)
						ResultArray(UBound(ResultArray)) = SelectedOption & ";" & "Not a valid property on target computer"
					Else
					
						If (IsArray(objItem.Properties_(SelectedOption).Value)) Then

							For Each Value in objItem.Properties_(SelectedOption).Value 
								ReDim Preserve ResultArray(UBound(ResultArray)+1)
								ResultArray(UBound(ResultArray)) = SelectedOption & ";" & Value
							Next
						Else
							ReDim Preserve ResultArray(UBound(ResultArray)+1)
							ResultArray(UBound(ResultArray)) = SelectedOption & ";" & CurrentProperty
						End If
					End If
					
				Next
			Next
		Work = ResultArray
	End Function
	
	Public Sub List_Options
		Dim Localhost_WMIService, Localhost_CollectionItems, Localhost_Properties
		
		Options = array()
		
		Set Localhost_WMIService = GetObject("winmgmts:\\localhost\root\cimv2")
		Set Localhost_CollectionItems = Localhost_WMIService.ExecQuery("Select * from " & WMI_Name)
		For Each objItem in Localhost_CollectionItems
		
			For Each Localhost_Properties In objItem.Properties_
				AddToArray = True
				
				' Check if the property is already added in the options array.
				For Each Option_Value in Options
					If (Option_Value = Localhost_Properties.Name) Then
						'Already in Options array
						AddToArray = False
					End If
				Next
				
				If (AddToArray) Then
					ReDim Preserve Options(UBound(Options)+1)
					Options(UBound(Options)) =  Localhost_Properties.Name
				End If
			Next
			
		Next	
	End Sub
End Class