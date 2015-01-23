'
' Each module needs to have:
' Class_Initialize, define the Name
' Work and List_Options Sub
'
'
Register_Class( New EventViewer)


Class EventViewer
	' Define Properties
	Public Name
	Public Options
	
	' Initialize the Class
	Private Sub Class_Initialize
		Name = "EventViewer"
		
	End Sub
	
	Public Function Work (ComputerName, SelectedOptionsArray)
		'MsgBox "[" & ComputerName & "]"
		
		Dim ResultArray, SubResult
		ResultArray = array()
		SubResult = array()
		
			Option_1 = SelectedOptionsArray(0)
			
			If (InStr(Option_1, ",") > 0) Then
				Option_1 = Split(SelectedOptionsArray(0),";")
				
				For Each Option_Log in Option_1
					
					SubResult = QueryWMI(ComputerName, Option_Log, SelectedOptionsArray(1))
					For Each SubResult_Index in SubResult
						ReDim Preserve ResultArray(UBound(ResultArray)+1)
						ResultArray(UBound(ResultArray)) = SubResult_Index
					Next ' Result Index For-Loop
					
				Next 'Option For-Loop
				
			Else
				'Option_1 = SelectedOptionsArray(0)
				SubResult = QueryWMI(ComputerName, Option_1 , SelectedOptionsArray(1))
				For Each SubResult_Index in SubResult
						ReDim Preserve ResultArray(UBound(ResultArray)+1)
						ResultArray(UBound(ResultArray)) = SubResult_Index
				Next
			End If
			
		Work = ResultArray
	End Function
	
	Private Function QueryWMI (ComputerName, LogName, Source)
		Dim WMIService, CollectionItems, SubResult
		SubResult = array()
		
		LogName_Temp = Split(LogName,";")
		LogName = LogName_Temp(1)
		
		Source_Temp = Split(Source,";")
		Source = Source_Temp(1)
		
		Set WMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & ComputerName & "\root\cimv2")
		If (InStr(Source, ",") > 0) Then
			Option_2 = Split(Source,";")
			
			For Each Option_Log in Option_2
			
				Set CollectionItems = WMIService.ExecQuery ("Select * from Win32_NTLogEvent Where Logfile = '" & LogName & "' and SourceName = '" & Option_Log & "'")	
				
				ReDim Preserve SubResult(UBound(SubResult)+1)
				MsgBox LogName & "/" & Option_Log & "; Count: " & CollectionItems.Count
				SubResult(UBound(SubResult)) = LogName & "/" & Option_Log & "; Count: " & CollectionItems.Count
				
			Next
		Else
			Option_2 = Source
				Set CollectionItems = WMIService.ExecQuery ("Select * from Win32_NTLogEvent Where Logfile = '" & LogName & "' and SourceName = '" & Option_2 & "'")
				
				ReDim Preserve SubResult(UBound(SubResult)+1)
				MsgBox LogName & "/" & Option_2 & ";Count: " & CollectionItems.Count
				SubResult(UBound(SubResult)) = LogName & "/" & Option_2 & ";Count: " & CollectionItems.Count
				
		End If
		
		QueryWMI = SubResult
	End Function
	
	Public Sub List_Options
		Dim Localhost_WMIService, Localhost_CollectionItems, Localhost_Properties
		
		Options = array()
		ReDim Preserve Options(UBound(Options)+1)
		Options(UBound(Options)) =  "Input;required;EventLog;Use comma(,) to seperate LogNames (Example: System, Applications)."
		ReDim Preserve Options(UBound(Options)+1)
		Options(UBound(Options)) =  "Input;required;EventSource;Use comma(,) to seperate Source (Example: Disk)"
		
	End Sub
End Class