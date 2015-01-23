'
' Each module needs to have:
' Class_Initialize, define the Name
' Work and List_Options Sub
'
'
Register_Class( New FileExists)


Class FileExists
	' Define Properties
	Public Name
	Public Options
	
	' Initialize the Class
	Private Sub Class_Initialize
		Name = "FileExists"
		
	End Sub
	
	Public Function Work (ComputerName, SelectedOptionsArray)
		'MsgBox "[" & ComputerName & "]"
		
		Dim ResultArray
		ResultArray = array()
			
		For Each SelectedOptions in SelectedOptionsArray
			Parameter = Split(SelectedOptions, ";")
			If (Parameter(0) = "Files") Then
				If (Len(Parameter(1))>0) Then
					If (InStr(Parameter(1), ",") > 0) Then
						FileNames = Split(Parameter(1),",")
						
						For Each FileName in FileNames
							
							ReDim Preserve ResultArray(UBound(ResultArray)+1)
							ResultArray(UBound(ResultArray)) = Test_FileExists(ComputerName, Parameter(0), FileName)
						Next
					Else
						
						ReDim Preserve ResultArray(UBound(ResultArray)+1)
						ResultArray(UBound(ResultArray)) = Test_FileExists(ComputerName, Parameter(0), Parameter(1))					
					End If
				End If ' End of length check
			End If
			
			If (Parameter(0) = "Folders") Then
				If (Len(Parameter(1))>0) Then
					If (InStr(Parameter(1), ",") > 0) Then
						FolderNames = Split(Parameter(1),",")
						
						For Each FolderName in FolderNames
							ReDim Preserve ResultArray(UBound(ResultArray)+1)
							ResultArray(UBound(ResultArray)) = Test_FolderExists(ComputerName, Parameter(0),  FolderName)
						Next
					Else
						ReDim Preserve ResultArray(UBound(ResultArray)+1)
						ResultArray(UBound(ResultArray)) = Test_FolderExists(ComputerName, Parameter(0), Parameter(1))
					End If	
				End If ' End of Length check
			End If
		Next
			
		Work = ResultArray
	End Function
	
	Private Function Test_FileExists (ComputerName, ParameterName, FileName)
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		If  Not (InStr(FileName,"\") = 1) Then
			FileName = "\" & FileName
		End If
		
		If (objFSO.FileExists("\\" & ComputerName & FileName)) Then
			Test_FileExists = ParameterName & ";" & "\\" & ComputerName & FileName & vbTab & "Exists"
		Else
			Test_FileExists = ParameterName & ";" & "\\" & ComputerName & FileName & vbTab & "Do not exist"
		End if
		
	End Function

	Private Function Test_FolderExists (ComputerName, ParameterName, FolderName)
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		If  Not (InStr(FolderName,"\") = 1) Then
			FolderName = "\" & FolderName
		End If
		
		If (objFSO.FolderExists("\\" & ComputerName & FolderName)) Then
			Test_FolderExists = ParameterName & ";" & "\\" & ComputerName & FolderName & vbTab & "Exists"
		Else
			
			Test_FolderExists = ParameterName & ";" & "\\" & ComputerName & FolderName & vbTab & "Do not exist"
		End if
		
	End Function
	
	Public Sub List_Options

		Options = array()
		ReDim Preserve Options(UBound(Options)+1)
		Options(UBound(Options)) =  "Input;Files;Use comma(,) to seperate filenames. Also, use network share (C$\Program Files\test.ini) paths, script appends string to \\comutername\."
		ReDim Preserve Options(UBound(Options)+1)
		Options(UBound(Options)) =  "Input;Folders;Use comma(,) to seperate folders. Also, use network share (C$\Program Files\Testfolder) paths, script appends string to \\comutername\"
		
	End Sub
End Class