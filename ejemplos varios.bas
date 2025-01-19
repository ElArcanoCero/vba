Attribute VB_Name = "ejemplo"
Sub suma()
Attribute suma.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Range("a12") = WorksheetFunction.Sum(Range("a2:a11"))
    Range("b12") = WorksheetFunction.Sum(Range("b2:b11"))
    
    For i = 2 To 11
        
        Range("c" & i).Value = Application.WorksheetFunction.Sum(Range("a" & i), Range("b" & i))
        
    Next i
    
    Range("c1") = "suma A, B"
    Range("c12") = WorksheetFunction.Sum(Range("c2:c11"))
        
End Sub

Sub porcentage()

    Range("d1") = "50% de A"
    Range("e1") = "80% de B"
    
    For i = 2 To 11
        
        Range("d" & i).Value = Range("a" & i) * 0.5
        Range("e" & i).Value = Range("b" & i) * 0.8
        
    Next i
    
End Sub

Sub media()
    
    Range("f1") = "media A+B"
    For i = 2 To 11
        
        Range("f" & i).Value = Range("c" & i) / 2
        
    Next i

End Sub
Sub centrar()
Attribute centrar.VB_ProcData.VB_Invoke_Func = " \n14"

    Range("a1").HorizontalAlignment = xlCenter
    Range("b1").HorizontalAlignment = xlCenter
    Range("c1").HorizontalAlignment = xlCenter
    Range("d1").HorizontalAlignment = xlCenter
    Range("e1").HorizontalAlignment = xlCenter
    Range("f1").HorizontalAlignment = xlCenter

End Sub

Sub color()
    
    Range("a1").Interior.color = RGB(255, 0, 0)
    Range("b1").Interior.color = RGB(0, 255, 0)
    Range("c1").Interior.color = RGB(0, 0, 255)
    Range("d1").Interior.color = RGB(125, 125, 0)
    Range("e1").Interior.color = RGB(0, 125, 125)
    Range("f1").Interior.color = RGB(125, 0, 125)
    Range("a2:f11").Interior.color = RGB(150, 150, 150)
    Range("a12:c12").Interior.color = RGB(0, 255, 0)
    
End Sub
Sub crear_grafico()

    Set Worksheet = ActiveWorkbook.Sheets("hoja1")
        Set ChartObject = Worksheet.ChartObjects.Add(Left:=550, Width:=350, Top:=15, Height:=200)
            ChartObject.Name = "grafico1"
            ChartObject.chart.ChartType = xlXYScatterLines
            ChartObject.chart.SetSourceData Source:=Worksheet.Range("a2:b11")
            ChartObject.chart.HasTitle = True
            ChartObject.chart.ChartTitle.Text = "regrecion"
            ChartObject.chart.ChartTitle.Font.Size = 20
            ChartObject.chart.ChartTitle.Font.Bold = True
            ChartObject.chart.ChartTitle.Font.color = RGB(255, 0, 0)
            ChartObject.chart.Axes(xlCategory).HasTitle = True
            ChartObject.chart.Axes(xlCategory).AxisTitle.Text = "datos X"
            ChartObject.chart.Axes(xlValue).HasTitle = True
            ChartObject.chart.Axes(xlValue).AxisTitle.Text = "datos Y"
            ChartObject.chart.HasLegend = falce
            
                Set series = ChartObject.chart.SeriesCollection(1)
                    Set trendline = series.Trendlines.Add(Type:=xlLinear)
                        trendline.DisplayEquation = True
                        trendline.DisplayRSquared = True
            
                               
End Sub

Sub Buscador()
Attribute Buscador.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim dia1 As Date
    Dim dia2 As Date
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim con As Long
    Dim p As Long
    
    x = 1
    y = 0
    p = 2
    con = 2
    dia1 = Date
    
    
    Range("A2").Select 'inicio '

    Do While ActiveCell <> Empty
    
        x = x + 1
        ActiveCell.Offset(1, 0).Select
          
    Loop
    
    Range("G1:Q" & x).ClearContents
    Columns("C").Delete
    
    For i = 2 To x
        
        z = StrComp(Range("A" & i), "IoT", 1)
        If z = 0 Then
            Range("A" & p).Value = Range("A" & i).Value
            Range("B" & p).Value = Range("B" & i).Value
            Range("C" & p).Value = Range("C" & i).Value
            Range("D" & p).Value = Range("D" & i).Value
            Range("E" & p).Value = Range("E" & i).Value
            Range("F" & p).Value = Range("F" & i).Value
            p = p + 1
                
        End If
         
    Next i
    
    For i = p To x
    
        Range("A" & i).ClearContents
        Range("B" & i).ClearContents
        Range("C" & i).ClearContents
        Range("D" & i).ClearContents
        Range("E" & i).ClearContents
        Range("F" & i).ClearContents
        
    Next i
      
    For i = 2 To p - 1
      
        dia2 = Range("C" & i).Value
        z = DateDiff("d", dia2, dia1)
        Range("F" & i) = z
        
        If z > 5 Then
        
            Range("H" & con).Value = Range("B" & i).Value
            Range("I" & con).Value = z
            con = con + 1
            
                            
        End If
                                                        
    Next i
    
    Range("I1").Sort Key1:=Range("I2"), Order1:=xlAscending, Header:=xlYes
    Range("G2").Value = dia1
    Range("G3").Value = p
    Range("G4").Value = con
    Columns("A:I").AutoFit
    Range("A1").Select
   
End Sub

Sub Buscador2()
Attribute Buscador2.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim p As Long
    
    x = 1
    y = 1
    z = 0
    p = 2
   
    
    Range("A2").Select 'inicio '

    Do While ActiveCell <> Empty
    
        x = x + 1
        ActiveCell.Offset(1, 0).Select
          
    Loop
    
    Range("B2").Select 'inicio '

    Do While ActiveCell <> Empty
    
        y = y + 1
        ActiveCell.Offset(1, 0).Select
          
    Loop
    
    For i = 2 To x
        
         For t = 2 To y
            
            z = StrComp(Range("A" & i), Range("B" & t), 1)
            If z = 0 Then
            
                Range("C" & p).Value = Range("A" & i).Value
                Range("B" & t).Interior.ColorIndex = 4
                p = p + 1
                
            End If
        
         
        Next t
         
    Next i
    
   
End Sub
Sub Buscador3()
Attribute Buscador3.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim dia1 As Date
    Dim dia2 As Date
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim con As Long
    Dim p As Long
    
    x = 1
    y = 0
    p = 2
    con = 2
    dia1 = Date
    
    
    Range("A2").Select 'inicio '

    Do While ActiveCell <> Empty
    
        x = x + 1
        ActiveCell.Offset(1, 0).Select
          
    Loop
    
    Range("G1:Q" & x).ClearContents
    Columns("C").Delete
    
    For i = 2 To x
        
        z = StrComp(Range("A" & i), "IoT", 1)
        If z = 0 Then
            Range("A" & p).Value = Range("A" & i).Value
            Range("B" & p).Value = Range("B" & i).Value
            Range("C" & p).Value = Range("C" & i).Value
            Range("D" & p).Value = Range("D" & i).Value
            Range("E" & p).Value = Range("E" & i).Value
            Range("F" & p).Value = Range("F" & i).Value
            p = p + 1
                
        End If
         
    Next i
    
    For i = p To x
    
        Range("A" & i).ClearContents
        Range("B" & i).ClearContents
        Range("C" & i).ClearContents
        Range("D" & i).ClearContents
        Range("E" & i).ClearContents
        Range("F" & i).ClearContents
        
    Next i
      
    For i = 2 To p - 1
      
        dia2 = Range("C" & i).Value
        z = DateDiff("d", dia2, dia1)
        Range("F" & i) = z
        
        If z > 5 Then
        
            Range("H" & con).Value = Range("B" & i).Value
            Range("I" & con).Value = z
            con = con + 1
            
                            
        End If
                                                        
    Next i
    
    Range("I1").Sort Key1:=Range("I2"), Order1:=xlAscending, Header:=xlYes
    Range("G2").Value = dia1
    Range("G3").Value = p
    Range("G4").Value = con
    Columns("A:I").AutoFit
    Range("A1").Select
   
End Sub

Sub Buscador4()
Attribute Buscador4.VB_ProcData.VB_Invoke_Func = "b\n14"
    
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim p As Long
    
    x = 2
    y = 2
    z = 0
    p = 2
   
    
    Range("A2").Select 'inicio '

    Do While ActiveCell <> Empty
    
        x = x + 1
        ActiveCell.Offset(1, 0).Select
          
    Loop
    
    Range("B2").Select 'inicio '

    Do While ActiveCell <> Empty
    
        y = y + 1
        ActiveCell.Offset(1, 0).Select
          
    Loop
    
    y = y - 1
    
    For i = 2 To x
    
        Range("A" & i).Interior.ColorIndex = 34
        
        For t = 2 To y
            
            Range("D" & p).Value = Range("A" & i).Value
            Range("D" & p).Interior.ColorIndex = 34
            
            z = StrComp(Range("A" & i), Range("B" & t), 1)
            If z = 0 Then
            
                Range("D" & p).Value = ""
                Range("D" & p).Interior.ColorIndex = 0
                Range("A" & i).Interior.ColorIndex = 0
                t = y
                              
            End If
        
         
        Next t
        If Range("D" & p).Value = "" Then
            
        p = p - 1
                              
        End If
        
        p = p + 1
        
    Next i
    
    Range("D" & p).Interior.ColorIndex = 0
    Range("A" & i - 1).Interior.ColorIndex = 0
    
    p = 2
    
    For i = 2 To y
    
        Range("B" & i).Interior.ColorIndex = 36
        
        For t = 2 To x
            
            Range("E" & p).Value = Range("B" & i).Value
            Range("E" & p).Interior.ColorIndex = 36
            
            z = StrComp(Range("B" & i), Range("A" & t), 1)
            If z = 0 Then
            
                Range("E" & p).Value = ""
                t = x
                              
            End If
        
         
        Next t
        If Range("E" & p).Value = "" Then
            
        p = p - 1
                              
        End If
        
        p = p + 1
        
    Next i
    
    p = 2
    
    For i = 2 To x
        
         For t = 2 To y
            
            z = StrComp(Range("A" & i), Range("B" & t), 1) 'cero se considera igual a ""'
            
            If z = 0 Then
            
                Range("C" & p).Value = Range("A" & i).Value
                Range("C" & p).Interior.ColorIndex = 17
                Range("B" & t).Interior.ColorIndex = 17
                Range("A" & i).Interior.ColorIndex = 17
                p = p + 1
                
            End If
        
         
        Next t
         
    Next i
    
    Range("A1").Select
    Range("F1").Interior.ColorIndex = 3
    Range("F2").Interior.ColorIndex = 4
    Range("F3").Interior.ColorIndex = 5
    Range("F4").Interior.ColorIndex = 6
    Range("F5").Interior.ColorIndex = 7
    Range("F6").Interior.ColorIndex = 8
    Range("F7").Interior.ColorIndex = 9
    Range("F8").Interior.ColorIndex = 10
    Range("F9").Interior.ColorIndex = 11
    Range("F10").Interior.ColorIndex = 12
    Range("F11").Interior.ColorIndex = 13
    Range("F12").Interior.ColorIndex = 14
    Range("F13").Interior.ColorIndex = 15
    Range("F14").Interior.ColorIndex = 16
    Range("F15").Interior.ColorIndex = 17
    Range("F16").Interior.ColorIndex = 18
    Range("F17").Interior.ColorIndex = 19
    Range("F18").Interior.ColorIndex = 20
    Range("F19").Interior.ColorIndex = 21
    Range("F20").Interior.ColorIndex = 22
    Range("F21").Interior.ColorIndex = 23
    Range("F22").Interior.ColorIndex = 24
    Range("F23").Interior.ColorIndex = 25
    Range("F24").Interior.ColorIndex = 26
    Range("G1").Interior.ColorIndex = 27
    Range("G2").Interior.ColorIndex = 28
    Range("G3").Interior.ColorIndex = 29
    Range("G4").Interior.ColorIndex = 30
    Range("G5").Interior.ColorIndex = 31
    Range("G6").Interior.ColorIndex = 32
    Range("G7").Interior.ColorIndex = 33
    Range("G8").Interior.ColorIndex = 34
    Range("G9").Interior.ColorIndex = 35
    Range("G10").Interior.ColorIndex = 36
    Range("G11").Interior.ColorIndex = 37
    Range("G12").Interior.ColorIndex = 38
    Range("G13").Interior.ColorIndex = 39
    Range("G14").Interior.ColorIndex = 40
    Range("G15").Interior.ColorIndex = 41
    Range("G16").Interior.ColorIndex = 42
    Range("G17").Interior.ColorIndex = 43
    Range("G18").Interior.ColorIndex = 44
    Range("G19").Interior.ColorIndex = 45
    Range("G20").Interior.ColorIndex = 46
    Range("G21").Interior.ColorIndex = 47
    Range("G22").Interior.ColorIndex = 48
    Range("G23").Interior.ColorIndex = 49
    Range("G24").Interior.ColorIndex = 50
    
   
End Sub
