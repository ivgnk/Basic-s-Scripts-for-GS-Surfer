Option Explicit
'#Uses "D:\GENIK\My\Tools.bas\SrfTypes.bas"
' ������ �� ������ GridDataArr_Krig.BAS
Sub Main
 Const NDat=40
 Const DatExt = ".dat"
 Const GrdExt = ".grd"
'Const CalcFN = "Calc.inf"

 Dim FDatN (1 To NDat)

 Dim Surf As Object
 Dim DirS As String, SN As String, DN  As String, FsN As String, FdN  As String, FInfN  As String, BlankFN  As String, FullBlankFN  As String, DGridFN As String
 Dim Xst, YSt, xmin1 As Double, xmax1 As Double, ymin1 As Double, ymax1 As Double
 Dim NCols As Long, NRows As Long, NT As Long, i As Long, NDatCurr As Long
 Dim IsBlank As Boolean
'------------------------------
 Debug.Clear
  DirS = "D:\Dog\2008\Suksun\Programms\Blank.bas\"

 FInfN = GetFilePath$(,"inf", DirS, "���� inf-�����",0)
 Debug.Print FInfN
'---------------------------
 Open FInfN For Input As #1
    Input #1,NDatCurr
    For i = 1 To NDatCurr
     Input #1,FDatN(i)
     Debug.Print FDatN(i)
    Next i
    Input #1,BlankFN
    Close #1
'---------------------------
 	On Error Resume Next 'Turn off error reporting.
	Set Surf = GetObject(,"Surfer.Application")
	If Err.Number<>0 Then
		Set Surf = CreateObject("Surfer.Application")
	End If
	On Error GoTo 0 'Turn on error reporting.
    If Surf.Windows.Count=0 Then Surf.Documents.Add(srfDocPlot)
	Surf.Visible = False
'---------------------------------
	 For i = 1 To NDatCurr
      SN=FDatN(i)
       Debug.Print i
 	   FsN=DirS+SN+GrdExt
       FullBlankFN=DirS+BlankFN
           DGridFN=DirS+SN+"B"+GrdExt
        Surf.GridBlank (InGrid := FsN, BlankFile := FullBlankFN,	OutGrid := DGridFN,  OutFmt := srfGridFmtS7)
       '-------------
     Next i
 Debug.Print "Done"
End Sub

