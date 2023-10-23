'Dim

Sub Main
Dim Selection As Object
Dim Doc As Object
Dim Plot As Object
Dim i As Integer
'--------------------------------
	Debug.Clear
	'Get existing Surfer instance, or create a new one If none exists.
	On Error Resume Next 'Turn off error reporting.
	Set surf = GetObject(,"Surfer.Application")
	If Err.Number<>0 Then
        Debug.Print "Not found Surfer"
		GoTo Prole
	End If
    surf.Visible = False

    'Set Doc = Surf.Documents
    Set Plot = Surf.Documents(1)
	If Err.Number<>0 Then
        Debug.Print "Not found Documents"
		GoTo Prole
	End If
    Set Selection = Plot.Selection
	If Err.Number<>0 Then
        Debug.Print "Not found Selection"
		GoTo Prole
	End If

    surf.Visible = True
  Prole:

 	Debug.Print "DONE --------------------------------------------------"
End Sub





	
