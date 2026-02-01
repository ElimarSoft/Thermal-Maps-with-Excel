Attribute VB_Name = "Ranges14"
Option Explicit

Const STEPS As Long = 10000

Type DataCell
    TextColor As Double
    BackColor As Double
    LBorder As Double
    TBorder As Double
    RBorder As Double
    BBorder As Double
End Type

Type DoublePair
    ind As Double
    val As Double
End Type

Type ConfigData
    DataColor As Double
    BoderColor(7) As DoublePair
End Type

Dim ConfigSh As Worksheet
Dim Cnf As ConfigData
Dim Dc() As DataCell

Const SheetName = "Data1"
Const tablename = "Table2"
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Public Sub Process()

    Config

    'Delete Shapes
    Dim ws1 As Worksheet
    Dim sh1 As Shape
    For Each ws1 In Worksheets
        If ws1.Visible Then
            For Each sh1 In ws1.Shapes
                If (Left(sh1.Chart.ChartTitle.Text, 5) = "TABLE") Then sh1.Delete
            Next
        End If
    Next ws1

    Dim name1 As name
    For Each name1 In ActiveWorkbook.Names
        If name1.RefersToRange.Worksheet.Visible Then
            If Left(name1.name, 5) = "TABLE" Then
                CalcTable (name1.name)
            End If
        End If
    Next

End Sub

Public Sub CalcTable(tablename As String)
    
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim vals As Double
    
    Dim Data1 As Variant
    Dim Data2 As Variant

    Dim flag1 As Boolean: flag1 = False
    
    Dim rg1 As Range
    Dim rg2 As Range
    
    Dim Cell As Range
    
    Set rg1 = Range(tablename)
    
'    For Each Cell In rg1
'        If Cell.Font.color <> Cnf.DataColor Then Cell.value = 0
'    Next
    
    Data1 = rg1
    Data2 = rg1
    
    Dim RowCount As Long: RowCount = rg1.Rows.count
    Dim ColCount As Long: ColCount = rg1.Columns.count

    ReDim Dc(1 To RowCount, 1 To ColCount)

    For i = LBound(Data1, 1) To UBound(Data1, 1)
        For j = LBound(Data1, 2) To UBound(Data1, 2)
            With rg1.Cells(i, j)
                Dc(i, j).TextColor = .Font.color
                Dc(i, j).BackColor = .Interior.color
                Dc(i, j).LBorder = ColorData(.Borders(xlEdgeLeft).color)
                Dc(i, j).TBorder = ColorData(.Borders(xlEdgeTop).color)
                Dc(i, j).RBorder = ColorData(.Borders(xlEdgeRight).color)
                Dc(i, j).BBorder = ColorData(.Borders(xlEdgeBottom).color)
            End With
            'If Not Dc(i, j).TextColor = Cnf.DataColor Then Data1(i, j) = 0
        Next j
    Next i

    For n = 1 To STEPS
          
         flag1 = Not flag1
         If flag1 Then
          
             For i = LBound(Data1, 1) + 1 To UBound(Data1, 1) - 1
                For j = LBound(Data1, 2) + 1 To UBound(Data1, 2) - 1
                    If Not Dc(i, j).TextColor = Cnf.DataColor Then
                        vals = 0
                        vals = vals + (Data1(i - 1, j) - Data1(i, j)) * Dc(i, j).TBorder
                        vals = vals + (Data1(i + 1, j) - Data1(i, j)) * Dc(i, j).BBorder
                        vals = vals + (Data1(i, j - 1) - Data1(i, j)) * Dc(i, j).LBorder
                        vals = vals + (Data1(i, j + 1) - Data1(i, j)) * Dc(i, j).RBorder
                        Data2(i, j) = Data1(i, j) + vals / 8
                    End If
                Next j
            Next i
        
        Else
        
            For i = LBound(Data2, 1) + 1 To UBound(Data2, 1) - 1
                For j = LBound(Data2, 2) + 1 To UBound(Data2, 2) - 1
                    If Not Dc(i, j).TextColor = Cnf.DataColor Then
                        vals = 0
                        vals = vals + (Data2(i - 1, j) - Data2(i, j)) * Dc(i, j).TBorder
                        vals = vals + (Data2(i + 1, j) - Data2(i, j)) * Dc(i, j).BBorder
                        vals = vals + (Data2(i, j - 1) - Data2(i, j)) * Dc(i, j).LBorder
                        vals = vals + (Data2(i, j + 1) - Data2(i, j)) * Dc(i, j).RBorder
                        Data1(i, j) = Data2(i, j) + vals / 8
                    End If
                Next j
            Next i
    
        End If
    
    Next n
    
    Dim r1 As Long: r1 = rg1.Rows(1).Row
    Dim r2 As Long: r2 = r1 + rg1.Rows.count - 1
    Dim c1 As Long: c1 = rg1.Columns(1).Column + rg1.Columns.count + 1
    Dim c2 As Long: c2 = c1 + rg1.Columns.count - 1
    
    With rg1.Worksheet
        Set rg2 = .Range(.Cells(r1, c1), .Cells(r2, c2))
    End With

    rg2 = Data1
    
    rg1.Copy
    rg2.PasteSpecial Paste:=xlPasteFormats
    
    DeleteRangeColorScale rg1
    ApplyColorScaleToTable rg2
    
    Chart rg2
    
    With rg1
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With rg2
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    
    'FillNotes tablename
    
End Sub

Private Sub FillNotes(tablename)

    Dim Cell As Range
    
    Range(tablename).ClearComments
    
    '    For Each Cell In Range(TableName)
    '        Dim Data As String
    '        Dim r1, c1 As Integer
    '        r1 = Cell.Row - Range(TableName).Row + 1
    '        c1 = Cell.Column - Range(TableName).Column + 1
    '
    '        Data = _
    '        CStr(Dc(r1, c1).LBorder) + ":" + _
    '        CStr(Dc(r1, c1).TBorder) + ":" + _
    '        CStr(Dc(r1, c1).RBorder) + ":" + _
    '        CStr(Dc(r1, c1).BBorder)
    '
    '        Cell.AddComment (Data)
    '    Next

End Sub

Private Function ColorData(col As Double) As Double

    Dim n As Integer

    ColorData = 1

    For n = 0 To 7

        If col = Cnf.BoderColor(n).ind Then
            ColorData = Cnf.BoderColor(n).val
        End If

    Next n

End Function

Private Sub Chart(rg1 As Range)
    
    Dim sh1 As Shape
    Set sh1 = rg1.Worksheet.Shapes.AddChart2()
    
    Dim r1 As Long: r1 = rg1.Rows(1).Row
    Dim c1 As Long: c1 = rg1.Columns(rg1.Columns.count).Column
    
    With sh1
        .Width = rg1.Width
        .Height = rg1.Height
        .Left = rg1.Worksheet.Cells(r1, c1 + 2).Left
        .Top = rg1.Top
    End With
    
    With sh1.Chart
        .SetSourceData Source:=rg1
        .ChartStyle = 307
        .ChartType = xlSurface
        .HasTitle = True
        .ChartTitle.Text = tablename
        .ChartArea.Shadow = True
        .ChartArea.Format.ThreeD.RotationX = -60
        .ChartArea.Format.ThreeD.RotationY = -15
    End With

End Sub

Private Sub Config()
    Set ConfigSh = Worksheets("Config")
    Cnf.DataColor = ConfigSh.Range("DataColor").Font.color
    
    Dim rg1 As Range
    Set rg1 = Range("BorderColors")
    Dim row1 As Range
    Dim n As Integer
    
    For Each row1 In rg1.Rows
        If row1.Cells(1, 2) <> vbNullString Then
            Cnf.BoderColor(n).ind = row1.Cells(1, 1).Borders(xlEdgeTop).color
            Cnf.BoderColor(n).val = row1.Cells(1, 2)
            n = n + 1
        End If
    Next

End Sub

Sub ApplyColorScaleToTable(rg1 As Range)

    Dim FormatColors(2) As Double
    FormatColors(0) = 8109667
    FormatColors(1) = 8711167
    FormatColors(2) = 7039480

    ' Añadir escala de 3 colores
    Dim cs As ColorScale
    Set cs = rg1.FormatConditions.AddColorScale(ColorScaleType:=3)

    ' Valor mínimo
    With cs.ColorScaleCriteria(1)
        .Type = xlConditionValueLowestValue
        .FormatColor.color = FormatColors(0)
        .FormatColor.TintAndShade = 0
    End With

    ' Percentil 50 (punto medio)
    With cs.ColorScaleCriteria(2)
        .Type = xlConditionValuePercentile
        .value = 50
        .FormatColor.color = FormatColors(1)
        .FormatColor.TintAndShade = 0
    End With

    ' Valor máximo
    With cs.ColorScaleCriteria(3)
        .Type = xlConditionValueHighestValue
        .FormatColor.color = FormatColors(2)
        .FormatColor.TintAndShade = 0
    End With

End Sub

Private Sub DeleteRangeColorScale(rg1 As Range)
    Dim fc As Object
    Dim i As Long

    ' Loop backwards to avoid skipping items when deleting
    For i = rg1.FormatConditions.count To 1 Step -1
        Set fc = rg1.FormatConditions(i)
        If fc.Type = xlColorScale Then fc.Delete
    Next i
End Sub

