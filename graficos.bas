Attribute VB_Name = "Module1"
Option Explicit

Sub crear_grafico()
Dim grafico As ChartObject
Dim wks As Worksheet
Set wks = ActiveWorkbook.Sheets(1)
Set grafico = wks.ChartObjects.Add(Left:=60, Width:=400, Top:=200, Height:=200)
grafico.Name = "Grafico_1"
grafico.Chart.ChartType = xlXYScatterLines
grafico.Chart.SetSourceData Source:=wks.Range("C7:M11")
End Sub

Sub seleccionar_grafico()
Dim wks As Worksheet
Set wks = ActiveWorkbook.Sheets(1)
wks.ChartObjects("Grafico_1").Delete
Selection.Delete
End Sub
Sub seleccionar_serie()
Dim cht As Chart
Dim wks As Worksheet
Set wks = ActiveWorkbook.Sheets(1)
Set cht = wks.ChartObjects("Grafico_1").Chart
On Error Resume Next
cht.SeriesCollection("Series1").Delete
On Error GoTo 0
End Sub

Sub borrar_series_1()
Dim cht As Chart
Dim i As Integer
Set cht = ActiveWorkbook.Sheets(1).ChartObjects("Grafico_1").Chart
For i = cht.SeriesCollection.Count To 1 Step -1
    cht.SeriesCollection(i).Delete
Next i
End Sub
Sub borrar_series_2()
Dim cht As Chart
Set cht = ActiveWorkbook.Sheets(1).ChartObjects("Grafico_1").Chart
Dim series As series
For Each series In cht.SeriesCollection
    series.Delete
Next series
End Sub

Sub añadir_datos()
Dim cht As Chart
Set cht = ActiveWorkbook.Sheets(1).ChartObjects("Grafico_1").Chart
cht.SeriesCollection.NewSeries.Select
With Selection
    .Name = Range("B8").Value
    .XValues = Range("C7:M7")
    .Values = Range("C8:M8")
End With
End Sub

Sub añadir_datos_2()
Dim cht As Chart
Dim i As Integer
Set cht = ActiveWorkbook.Sheets(1).ChartObjects("Grafico_1").Chart
For i = 1 To 4
    cht.SeriesCollection.NewSeries.Select
    With Selection
        .Name = Cells(i + 7, 2).Value
        .XValues = Range("C7:M7")
        .Values = Range("C8:M8").Offset(i - 1, 0)
    End With
Next i
End Sub

Sub editar_titulo()
Dim cht As Chart
Set cht = ActiveWorkbook.Sheets(1).ChartObjects("Grafico_1").Chart
cht.HasTitle = True
cht.ChartTitle.Text = "Grafico 1"
With cht.ChartTitle.Font
    .Size = 16
    .Bold = True
    .Color = RGB(255, 0, 0)
End With
End Sub

Sub editar_ejes()
Dim cht As Chart
Set cht = ActiveWorkbook.Sheets(1).ChartObjects("Grafico_1").Chart
On Error Resume Next
cht.Axes(xlCategory, xlPrimary).Delete
cht.Axes(xlValue, xlPrimary).Delete
On Error GoTo 0
On Error Resume Next
cht.HasAxis(xlCategory, xlPrimary) = True
cht.HasAxis(xlValue, xlPrimary) = True
On Error GoTo 0
cht.Axes(xlCategory).HasMajorGridlines = True
cht.Axes(xlCategory).HasMinorGridlines = True
cht.Axes(xlValue).HasMajorGridlines = True
cht.Axes(xlValue).HasMinorGridlines = True
cht.Axes(xlCategory).HasTitle = True
cht.Axes(xlCategory).AxisTitle.Characters.Text = "Año"
cht.Axes(xlCategory).MaximumScale = 2020
cht.Axes(xlCategory).MinimumScale = 2000
cht.Axes(xlCategory).MajorUnit = 2
cht.Axes(xlCategory).MinorUnit = 1
cht.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
cht.Axes(xlCategory).MinorTickMark = xlTickMarkOutside
cht.Axes(xlCategory).ReversePlotOrder = True
cht.Axes(xlCategory).LogBase = 10
cht.Axes(xlCategory).DisplayUnit = xlHundreds
cht.Axes(xlCategory).TickLabelPosition = xlTickLabelPositionNextToAxis
cht.Axes(xlCategory).Crosses = xlMaximum
End Sub

Sub modificar_serie()
Dim cht As Chart
Dim srs As series
Set cht = ActiveWorkbook.Sheets(1).ChartObjects("Grafico_1").Chart

cht.HasLegend = True
cht.Legend.Font.Color = RGB(150, 150, 0)
cht.Legend.Position = xlLegendPositionBottom
cht.Legend.Format.Fill.ForeColor.RGB = RGB(50, 50, 0)

Set srs = cht.SeriesCollection(1)
srs.Format.Line.ForeColor.RGB = RGB(255, 0, 0)
srs.Format.Line.Weight = xlThin
srs.Format.Line.Style = msoLineThinThick
srs.Format.Line.DashStyle = msoLineDashDot
srs.Format.Line.Transparency = 0.5
srs.MarkerSize = 10
srs.MarkerStyle = xlMarkerStyleDiamond
srs.MarkerBackgroundColor = RGB(0, 255, 0)
srs.MarkerForegroundColor = RGB(0, 0, 255)

End Sub
