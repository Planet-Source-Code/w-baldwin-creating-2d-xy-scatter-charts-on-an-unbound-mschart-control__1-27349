<div align="center">

## Creating 2D XY scatter charts on an unbound MSChart Control


</div>

### Description

Creating bar and line charts is fairly well documented on the MSDN CD, but 2D XY (what I call XY scatter graphs) are not. This example demonstrates a method to not only create 2D XY graphs, but also plot multiple series on the same chart. The things that need to be done are pretty esoteric, including using multiple DataGrid columns for each series, removing null points from one or more series, etc., which makes this kind of graph very unique.

This example also illustrates some other properties and methods that can be useful for other types of charts.
 
### More Info
 
To use the example:

1) Open a new project

2) Add the MSChart component to your toolbox

3) Add an MSChart control to your form. For best results, resize it to at least 8 inches wide by 5 inches tall.

4) Paste the code into the code window, and

5) Run


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[W\. Baldwin](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/w-baldwin.md)
**Level**          |Intermediate
**User Rating**    |4.6 (37 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/w-baldwin-creating-2d-xy-scatter-charts-on-an-unbound-mschart-control__1-27349/archive/master.zip)

### API Declarations

```
Option Explicit
'This example shows how to plot multiple X-Y scatter graphs, also
'known as 2D XY, on an unbound MS Chart control. X-Y scatter graphs
'differ from other types because the associated DataGrid object needs
'2 columns per series, rather than just one. The first column for
'each series stores the X values, and the second one stores the Y
'values. Another difference is that if the # of plot points differs
'between multiple series, you have to remove the null points of the
'shorter series.
'Instructions: Add the MSChart control to your toolbox, then add
'it to a form. For clarity, try to make it at least 8 inches
'wide, by 5 inches tall, on the form. then, paste this code into
'the code window, and run.
'written by W. Baldwin, 8/2001
'OldRowCount keeps track of how many points have been plotted for
'the previous series, so we can remove null points from all the
'series that are shorter:
Dim OldRowCount As Long
'PenColor determines whether we are drawing in color or Black & White
'Black & White is for black and white printers, and uses different
'line patterns to distinguish series.
'ShowMarker is a flag that determines whether each plot point
'has a marker or not.
Dim PenColor As Boolean, ShowMarker As Boolean
'ChartPoints is the array that will hold the plot data
Dim ChartPoints() As Double
Dim lRow As Long, lRow2 As Long
Dim i As Integer, MsgPrompt As String
Dim XValue As Single, YValue As Single
```


### Source Code

```
Private Sub Form_Load()
With MSChart1
  .chartType = VtChChartType2dXY
  .ShowLegend = True
  With .Plot.Axis(VtChAxisIdY).AxisTitle
    .VtFont.Size = 12
    .Visible = True
    .Text = "Y Axis text"
  End With
  With .Plot.Axis(VtChAxisIdX).AxisTitle
    .VtFont.Size = 12
    .Visible = True
    .Text = "X Axis text"
  End With
  .Title.VtFont.Size = 12
  .Title = "Example 2D XY Scatter Graph"
  .Legend.Location.LocationType = VtChLocationTypeBottom
  .Plot.Axis(VtChAxisIdY).AxisScale.Type = VtChScaleTypeLinear
  .Plot.Axis(VtChAxisIdX).AxisScale.Type = VtChScaleTypeLinear
  'Tip from KB article Q194221:
  .Plot.UniformAxis = False
  .Footnote.Text = "Footnote goes here"
End With
PenColor = True 'Draw in color
ShowMarker = True 'Show plot points
ChartIt 1
PenColor = False 'Black and White
ShowMarker = False 'Don't show plot points
ChartIt 2
End Sub
Private Sub ChartIt(CurSeries As Integer)
MousePointer = 11
'Create a new array of plot points for this Series
'We will redim the first subscript differently, to show that each
'series can have a different # of plot points:
If CurSeries = 1 Then
  ReDim ChartPoints(1 To 15, 1 To 2)
Else
  ReDim ChartPoints(1 To 10, 1 To 2)
End If
'Create the array data:
For lRow = 1 To UBound(ChartPoints, 1)
  'create the X and Y values:
  XValue = lRow + Rnd * 2
  YValue = lRow + Rnd * 2
  'create negative values for the 2nd series:
  If CurSeries = 2 Then
    XValue = XValue * -1
    YValue = YValue * -1
  End If
  ChartPoints(lRow, 1) = XValue
  ChartPoints(lRow, 2) = YValue
Next lRow
'We need to increase the ColumnCount. For X-Y Scatter graphs, we
'need 2 columns for each series.
MSChart1.ColumnCount = CurSeries * 2
With MSChart1
  With .Plot
    .Wall.Brush.Style = VtBrushStyleSolid
    'Normally, you might want the Wall background of the Chart
    'to be in color, if you're using Color pens, and to be white
    'if using B&W pens, but, since we're drawing both a color
    'series *and* a B&W series on *one* chart, we'll just make
    'the wall color, for now. If you want White, uncomment the
    'line found about 10 lines down:
    .Wall.Brush.FillColor.Set 255, 255, 225
    If PenColor Then
      .Wall.Brush.FillColor.Set 255, 255, 225
      'You can set the individual Pen colors here, or just use
      'the defaults.
    Else 'Based on an article in the VB KB:
      'Uncomment the next line if you want the wall color to
      'be white:
      '.Wall.Brush.FillColor.Set 255, 255, 255
      'Set the different patterns for Black and White plotting.
      'You need to set the Pen for only the 'X' column:
      Select Case CurSeries * 2 - 1
      Case 1
       .SeriesCollection(1).Pen.Style = VtPenStyleSolid
       .SeriesCollection(1).Pen.VtColor.Set 0, 0, 0
      Case 3
       .SeriesCollection(3).Pen.Style = VtPenStyleDashed
       .SeriesCollection(3).Pen.VtColor.Set 0, 0, 0
      Case 5
       .SeriesCollection(5).Pen.Style = VtPenStyleDotted
       .SeriesCollection(5).Pen.VtColor.Set 0, 0, 0
      Case 7
       .SeriesCollection(7).Pen.Style = VtPenStyleDitted
       .SeriesCollection(7).Pen.VtColor.Set 0, 0, 0
      End Select
    End If
  End With
  .ColumnLabelCount = CurSeries * 2
  'If the current series has more plot points that the previous
  'one, we need to change .RowCount accordingly:
  If UBound(ChartPoints, 1) > OldRowCount& Then
    .RowCount = UBound(ChartPoints, 1)
  End If
  'Both of the next 2 lines seem to do the same thing:
  .Plot.SeriesCollection(CurSeries * 2 - 1).SeriesMarker.Show = ShowMarker
  .Plot.SeriesCollection.Item(CurSeries * 2 - 1).SeriesMarker.Show = ShowMarker
  'Create the plot points for this series from the ChartPoints array:
  For lRow = 1 To UBound(ChartPoints, 1)
     .DataGrid.SetData lRow, CurSeries * 2 - 1, ChartPoints(lRow, 1), False
     .DataGrid.SetData lRow, CurSeries * 2, ChartPoints(lRow, 2), False
  Next
  'Remove null points from *this* series, if it has *fewer*
  'points than the prior ones. If you don't remove null points,
  'then the graph will add 0,0 points, erroneously. See MS
  'Knowledge Base article Q177685 for more info:
  For lRow2 = lRow To OldRowCount&
    .DataGrid.SetData lRow2, CurSeries * 2 - 1, 0, True
    .DataGrid.SetData lRow2, CurSeries * 2, 0, True
  Next
  'Remove null points from *prior* series, if this series
  'has *more* points than the prior ones:
  If CurSeries > 1 Then
    For lRow = OldRowCount& + 1 To .RowCount
      For lRow2 = 1 To CurSeries - 1
        .DataGrid.SetData lRow, lRow2 * 2 - 1, 0, True
        .DataGrid.SetData lRow, lRow2 * 2, 0, True
      Next
    Next
  End If
  'Store the current RowCount
  OldRowCount& = .RowCount
  .Column = CurSeries * 2 - 1
  .ColumnLabel = "Series " & Str(CurSeries)
  .Refresh
End With
SubExit:
MousePointer = 0
End Sub
Private Sub Command1_Click()
'If you want to be able to print the chart, just
'add a Command1 button that calls this event.
MsgPrompt = "Make sure that device " & Printer.DeviceName & " is ready"
i = MsgBox(MsgPrompt, vbOKCancel, "Confirmation")
If i = vbCancel Then
  Exit Sub
End If
On Error GoTo ErrHandler
MSChart1.EditCopy
Printer.Print " "
Printer.PaintPicture Clipboard.GetData(), 0, 0
Printer.EndDoc
Exit Sub
ErrHandler:
  i = MsgBox("An error has occurred. Make sure your selected printer can print graphics.", vbOKOnly, "Error")
  Resume Next
End Sub
```

