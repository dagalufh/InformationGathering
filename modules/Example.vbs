'
' Each module needs to have:
' Class_Initialize, define the Name
' Work and List_Options Sub
'
'
Register_Class( New Example)


Class Example
	' Define Properties
	Public Name
	Public Options

	' Initialize the Class
	Private Sub Class_Initialize
		Name = "Example"
		
	End Sub
	
	Public Function Work (ComputerName, SelectedOptionsArray)
		'MsgBox "[" & ComputerName & "]"
		Dim ResultArray
		ResultArray = array()
		' Do stuff with the SelectedOptionsArray
		For Each SelectedOption in SelectedOptionsArray
		
			' Currently we do not to anything with the values returned. This is to show you what they are. Run it, and check the generated logfile.
			ReDim Preserve ResultArray(UBound(ResultArray)+1)
			ResultArray(UBound(ResultArray)) = SelectedOption
		Next

		Work = ResultArray
	End Function
	
	Public Sub List_Options
		Dim Localhost_WMIService, Localhost_CollectionItems, Localhost_Properties
		
		Options = array()
		ReDim Preserve Options(UBound(Options)+1)
		Options(UBound(Options)) =  "RegularProperty"
		ReDim Preserve Options(UBound(Options)+1)
		Options(UBound(Options)) =  "Input;Required;TextProperty;This is a description. This input text box is required if this module is selected."
		ReDim Preserve Options(UBound(Options)+1)
		Options(UBound(Options)) =  "Input;OptionalProperty;This is also a description, this input text box is optional."
		
		
	End Sub
End Class