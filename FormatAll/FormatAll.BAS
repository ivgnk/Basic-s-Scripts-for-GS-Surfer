
Dim S As String
'----------------
Dim SurferApp As Object
Dim Sh As Object
Dim Win As Object
Dim Plot As Object
'----------------
Dim Overlays As Object
Dim ContourMap As Object
Dim ContourLevels As Object
Dim BaseMap As Object
Dim Post As Object
'----------------
Dim Axes As Object
Dim ScaleBar As Object
Dim ScaleBars As Object
Dim FontFormat As Object
Sub Main
 Debug.Print "----------------------------------------"
 Debug.Print "srfPostPosCenter	="+Str(srfPostPosCenter)
 Debug.Print "srfPostPosLeft	="+Str(srfPostPosLeft)
 Debug.Print "srfPostPosRight	="+Str(srfPostPosRight)
 Debug.Print "srfPostPosAbove	="+Str(srfPostPosAbove)
 Debug.Print "srfPostPosBelow	="+Str(srfPostPosBelow)
 Debug.Print "srfPostPosUser	="+Str(srfPostPosUser)
End Sub

Sub ssMain
'#Uses "d:\Genik\bi\Work\SurfScr\UStrings.BAS"
' Как форматируется каждый элемент
Dim i As Integer
Dim j As Integer
'========================
Debug.Print "--------------------------"
Set SurferApp = CreateObject("Surfer.Application")
 'Makes Surfer unvisible
SurferApp.Visible = False
 'Assigns the Documents collection to the variable named "Docs"
Set Docs = SurferApp.Documents
 '---------------Replacing
Set Plot = Docs.Open(FileName:="D:\1\3_12.srf")
Set Sh = SurferApp.Documents(1).Shapes
'Debug.Print Sh.Count
For I=1 To Sh.Count
 'Debug.Print SWLZ(i,2)+"  "+Sh.Item(Index:=i)
 Select Case UCase(Sh.Item(i))
  Case "TEXT"
   Call Format_Text(i)
  Case "POLYLINE"   
   Format_Polyline(i)
  Case "RECTANGLE"   
   Format_Rectangle(i)
  Case "POLYGON"   
   Format_Polygon(i)
  Case "SYMBOL"   
   Format_Symbol(Sh(i))
   
  Case "COMPOSITE"   
   'Sh(i).BreakApart
   'Debug.Print "----------------AfterBreaking"
   'Debug.Print Sh.Count
   'For j=1 To Sh.Count  
   ' Debug.Print SWLZ(j,2)+"  "+Sh.Item(Index:=j)
   'Next 
  Case "MAP"
   Format_Map(Sh(i))  
 End Select 
Next
SurferApp.Visible = True
End Sub
Sub TheFormat_Font(TheFont,Color_,Face_,Size_,Bold_)
  TheFont.Color= Color_
  TheFont.Face = Face_
  TheFont.Size = Size_
  TheFont.Bold = Bold_
End Sub
Sub TheFormat_Line(TheLine,Color_,Style_,Width_)
'---------Форматирование Line
 TheLine.ForeColor=Color_
 TheLine.Style    =Style_
 TheLine.Width    =Width_
End Sub
Sub TheFormat_Region(TheFill,BackColor_,ForeColor_,Pattern_,Transparent_)
'---------Форматирование Line
 TheFill.BackColor=BackColor_
 TheFill.ForeColor=ForeColor_
 'Pattern Styles 
'------------ 
'- - - With Transparencies
'None             'Solid
'Diagonal Cross   'Crosshatch
'Forward Slash    'Backward Slash
'Vertical         'Horizontal
'- - - Without Transparencies
'6.25% Black      '12.5% Black 
'25.0% Black      '50.0% Black 
'75.0% Black      '87.5% Black
'Sand             'Swamp 
'Forest           'Orchard 
'Water            'Limestone 
'Tilted Limestone 'DoАфсу=lomite 
'Shale            'Granite 
'Igneous          'Schist 
'Gravel           'Iron 
'Spheres          'Solid Dash 
'Thatch           'Checkerboard
'Shingle
'------------ 
 TheFill.Pattern=Pattern_
 TheFill.Transparent = Transparent_
End Sub
Sub TheFormat_Symbol(TheSymbol, Color_, Index_, Set_, Size_)
 TheSymbol.Color = Color_
 TheSymbol.index = Index_
 TheSymbol.Set   = Set_
 TheSymbol.Size  = Size_
 'TheSymbol.Select
End Sub
Sub Format_Text(i)
 Sh(i).Text=Replace(Sh(i).Text,"3.","4.")
' Call TheFormat_Font(Sh(i).Font,RGB(0,0,0),"Arial Cyr",True)
End Sub

Sub Format_Polyline(i)
' Sh(i).Font.Bold=True
' Sh(i).Font.Color=RGB(0,0,0)
' Sh(i).Font.Face="Arial Cyr"
' Sh(i).Line.ForeColor=RGB(255,0,0)
' Sh(i).Line.Style=".1 In. Dash"
' Sh(i).Line.Width=0.1 ' 0.1 cm
'------------ 
 Call TheFormat_Line(Sh(i).Line,RGB(255,255,0),".1 In. Dash",0.1)
End Sub
Sub Format_Rectangle(i)
 Call TheFormat_Line(Sh(i).Line,RGB(255,0,0),".1 In. Dash",0.1)
 'Sh(i).Line.ForeColor=RGB(255,0,0)
 'Sh(i).Line.Style=".1 In. Dash"
 'Sh(i).Line.Width=0.1 ' 0.1 cm
 Call TheFormat_Region(Sh(i).Fill,RGB(255,255,0),RGB(0,255,0),"Vertical",True)
 'Sh(i).Fill.BackColor=RGB(255,255,0)
 'Sh(i).Fill.ForeColor=RGB(0,255,0)
 'Sh(i).Fill.Pattern="Vertical"
 'Sh(i).Fill.Transparent = True
End Sub
Sub Format_Polygon(i)
'---------Форматирование тоже, что и Format_Rectangle(i)
 Call TheFormat_Line(Sh(i).Line,RGB(255,0,0),".1 In. Dash",0.1)
 Call TheFormat_Region(Sh(i).Fill,RGB(255,255,0),RGB(0,255,0),"Vertical",True)
End Sub

Sub Format_Symbol(TheSymbol)
  Call TheFormat_Symbol(TheSymbol.Marker, RGB(255,0,0), 9, "Mapinfo Cartographic", 0.38)
' TheSymbol.Marker.Color=RGB(255,0,0)
' TheSymbol.Marker.index=9
' TheSymbol.Marker.Set="Mapinfo Cartographic"
' TheSymbol.Marker.Size=0.38 '0.38 cm - in map units
' TheSymbol.Select
End Sub
Sub Format_Map(TheMapFrame)
Dim TAxisName As String
  Debug.Print "----------------MAP-Overlays
   Set Overlays = TheMapFrame.Overlays
   For J=1 To Overlays.Count
    Debug.Print SWLZ(j,2)+"  "+Overlays.Item(Index:=j)
    Select Case UCase(Overlays.Item(Index:=j))
     Case "CONTOURS","DEPTHV"
      Call Format_ContourMap(Overlays(J))
     Case "BASE","DEPOSIT","BOUNDARY","ATOL","LINE"
      Call Format_BaseMap(Overlays(J))
     Case "POST","NAME","TOWNS"
      Call Format_PostMap(Overlays(J))
     Case "CLASSED POST"
      Call Format_ClassedPostMap(Overlays(J))
    End Select
   Next  
 '------------------------------------------------------
  'Debug.Print "----------------MAP-Axes
   Set Axes = TheMapFrame.Axes
   For J=1 To Axes.Count
    TAxisName=Axes.Item(Index:=j)
    'Debug.Print SWLZ(j,2)+"  "+TAxisName
    Format_Axes Axes(J), TAxisName
   Next  
 '------------------------------------------------------
  'Debug.Print "----------------MAP-Scale Bar
   Set ScaleBars = TheMapFrame.ScaleBars
   'Debug.Print "Всего линеек  "+Str(ScaleBars.Count)
   For J=1 To ScaleBars.Count
    'Debug.Print SWLZ(j,2)+"  "+ScaleBars.Item(Index:=j)
    'Format_Axes Axes(J), TAxisName
    Format_ScaleBar(ScaleBars(J))  
   Next  
  'Debug.Print "-Different MAPs: Base, Contour, Post
  
End Sub

Sub Format_Axes(TheAxis,TheAxisName)
 'Debug.Print "The Title is "+TheAxisName
 Select Case TheAxisName 
'------Axis With Labels    
  Case "Bottom Axis" 
   TheAxis.Visible=True
   'TheAxis.ShowLabels=True
   'TheAxis.MajorTickType=srfTickOut
   'TheAxis.MinorTickType=srfTickOut
   'TheAxis.MinorTicksPerMajor = 4
   'TheAxis.AutoScale=False
   'Object.SetScale(Minimum, Maximum, _
   'MajorInterval, _
   'FirstMajorTick, LastMajorTick, Cross1, Cross2 )
   TheAxis.SetScale(440,560,20,440,560,504)
   '--------см. Text
   'TheAxis.LabelFont.Face="Arial Cyr"
   'TheAxis.LabelFont.Color=RGB(0,0,0)
   'TheAxis.Minimum=440
   'TheAxis.Maximum=560
   'TheAxis.Cross1=504
   'TheAxis.FirstMajorTick=440
   'TheAxis.LastMajorTick =560
  Case "Left Axis" 
   TheAxis.Visible=True
   'TheAxis.ShowLabels=True
   'TheAxis.MajorTickType=srfTickOut
   'TheAxis.MinorTickType=srfTickOut
   'TheAxis.MinorTicksPerMajor = 4
   'TheAxis.AutoScale=False
   'TheAxis.SetScale(504,740,20,520,740,440)
   '--------см. Text
   'TheAxis.LabelFont.Face="Arial Cyr"
   'TheAxis.LabelFont.Color=RGB(0,0,0)
   'TheAxis.Minimum=504
   'TheAxis.Maximum=740
   'TheAxis.Cross1=440
   'TheAxis.FirstMajorTick=520
   'TheAxis.LastMajorTick =740
   
'------Axis Without Labels   
  Case "Top Axis" 
   TheAxis.Visible=True
   TheAxis.ShowLabels=False
   TheAxis.MajorTickType=srfTickNone
   TheAxis.MinorTickType=srfTickNone
   TheAxis.AutoScale=True
   
  Case "Right Axis" 
   TheAxis.Visible=True
   TheAxis.ShowLabels=False
   TheAxis.MajorTickType=srfTickNone
   TheAxis.MinorTickType=srfTickNone
   TheAxis.AutoScale=True
   
  End Select 
End Sub

Sub Format_Scalebar(TheScalebar)
 'TheScalebar.Visible=False
 TheScalebar.NumCycles=4
 TheScalebar.CycleSpacing=10
 TheScaleBar.LabelIncrement=10
 TheScaleBar.LabelFont.Face ="Arial"
 TheScaleBar.LabelFont.Color=RGB(0,0,0)
 TheScaleBar.LabelFont.Size =6
End Sub

Sub Format_ContourMap(TheContourMap)
 TheContourMap.FillContours   = False
 TheContourMap.ShowColorScale = False
 TheContourMap.SmoothContours = srfConSmoothNone
 'TheContourMap.ColorScale.LabelFont.Face = "Times New Roman"
 'TheContourMap.GridFile = SurferApp.Path+"\Samples\demogrid.grd"
 
 Set ContourLevels = TheContourMap.Levels
 'ContourLevels.LoadFile("D:\1.lvl")
 'ContourLevels.SaveFile("D:\1.lvl")
 '------------------------------
 'Dim CustomLevels() As Double
 'CustomLevels(0) = 0         'CustomLevels(1) = 17
 'CustomLevels(2) = 43        'CustomLevels(3) = 56
 'CustomLevels(4) = 87        'CustomLevels(5) = 100
 'Levels.Set(Levels:=CustomLevels)
 '------------------------------
 'ContourLevels.AutoGenerate(MinLevel:=0, MaxLevel:=100, Interval:=10)
'--------------------------------
 TheContourMap.LabelEdgeDist=0.1 '0.1 cm
 TheContourMap.LabelFont.Face="Arial"
 TheContourMap.LabelFont.Color=RGB(0,0,0)
 TheContourMap.LabelFont.Size=8
 TheContourMap.LabelFont.Bold=False
 TheContourMap.LabelLabelDist = 5.08
 TheContourMap.LabelTolerance = 1.015
 ContourLevels.SetLabelFrequency(FirstIndex:=1, NumberToSet:=1, NumberToSkip:=0)
End Sub

Sub Format_BaseMap(TheBaseMap)
 'Debug.Print TheBaseMap.FileName
 Call TheFormat_Line  (TheBaseMap.Line,RGB(255,0,0),".1 In. Dash",0.1)
 Call TheFormat_Region(TheBaseMap.Fill,RGB(255,255,0),RGB(0,255,0),"Vertical",True)
 Call TheFormat_Symbol(TheBaseMap.Symbol, RGB(255,0,0), 9, "Mapinfo Cartographic", 0.1)
' TheBaseMap.xMin=400
End Sub

Sub Format_PostMap(ThePostMap)
 'ThePostMap.DataFile=SurferApp.Path+"\Samples\dgrid1.grd"
 'ThePostMap.xCol=1
 'ThePostMap.yCol=2
 'ThePostMap.LabCol = 3 '<--Column with label
 ThePostMap.SymFrequency = 1
'----------------------
'- SrfPostPosType Values
'- Value	Description
'- srfPostPosCenter	labels centered On the symbols 
'- srfPostPosLeft	labels To the Left of the symbols 
'- srfPostPosRight	labels To the Right of the symbols 
'- srfPostPosAbove	labels above the symbols
'- srfPostPosBelow	labels below the symbols 
'- srfPostPosUser	labels offset by user defined amounts 
 ThePostMap.LabelPos=srfPostPosAbove
 Call TheFormat_Symbol(ThePostMap.Symbol, RGB(255,0,0), 9, "Mapinfo Cartographic", 0.1)
 Call TheFormat_Font  (ThePostMap.LabelFont,RGB(0,0,255),"Times New Roman Cyr",10,False)
End Sub

Sub Format_ClassedPostMap(TheClassedPostMap)
End Sub
