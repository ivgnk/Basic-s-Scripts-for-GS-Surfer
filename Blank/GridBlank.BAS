Sub Main
Const NGrd=11
Const NBln=3
Dim SMainDir,DMainDir,SExt,DExt, SGrid, SBln, DGrid, S As String
Dim FGrdN (1 To NGrd)
Dim FBlnN (1 To NBln)
Dim PrfxBbln (0 To NBln)

Dim i,j,k As Integer
'--------------------------------
    GrdExt=".grd"
    BlnExt=".bln"
    FGrdN(1) ="Rlf"
    FGrdN(2) ="AkS"
    FGrdN(3) ="Ia"
    FGrdN(4) ="Ik"
    FGrdN(5) ="Ip"
    FGrdN(6) ="IIk"
    FGrdN(7) ="IIp"
    FGrdN(8) ="III"
    FGrdN(9) ="VII"
    FGrdN(10)="VIII"
    FGrdN(11)="F"
'---------------------------------
    FBlnN(1) ="Okont1n"
    FBlnN(2) ="Okont2nV"
    FBlnN(3) ="Okont3nV"
'---------------------------------
    PrfxBbln(0) =""
    PrfxBbln(1) ="B"
    PrfxBbln(2) ="B2"
    PrfxBbln(3) ="B3"
'---------------------------------
    SMainDir="D:\GENIK\MY\TOOLS.GVM\Delphi\DVol\USD\"
    DMainDir=SMainDir
    Debug.Clear
	'Get existing Surfer instance, or create a new one If none exists.
	On Error Resume Next 'Turn off error reporting.
	Set surf = GetObject(,"Surfer.Application")
	If Err.Number<>0 Then
		Set surf = CreateObject("Surfer.Application")
	End If
	On Error GoTo 0 'Turn on error reporting.
    If surf.Windows.Count=0 Then surf.Documents.Add(srfDocPlot)
	surf.Visible = False
'---------------------------------
	For i=1 To NBln
     Debug.Print i
	 For j=1 To NGrd
      S=FGrdN(j)
 	  Debug.Print S
 	   SGrid=SMainDir+FGrdN(j)+PrfxBbln(i-1)+GrdExt
 	   SBln =SMainDir+FBlnN(i)+BlnExt
       DGrid=SMainDir+FGrdN(j)+PrfxBbln(i)  +GrdExt
surf.GridBlank (InGrid := SGrid, BlankFile := SBln,	OutGrid := DGrid, OutFmt := srfGridFmtS7)
		'srfGridFmtBinary,srfGridFmtAscii,srfGridFmtS7,srfGridFmtXYZ
     Next j
    Next i

 	Debug.Print "DONE --------------------------------------------------"
    surf.Quit
End Sub





	
