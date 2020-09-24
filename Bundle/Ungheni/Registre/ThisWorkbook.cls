VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 4 To lastRow
        For j = 1 To 14
            Cells(i, j).VerticalAlignment = xlVAlignCenter
            If Len(Cells(i, j)) <= 3 Then
                Cells(i, j).HorizontalAlignment = xlHAlignCenter
            End If
        Next j
        Cells(i, 13).VerticalAlignment = xlVAlignCenter
        Cells(i, 13).HorizontalAlignment = xlHAlignCenter
        Cells(i, 13).Font.Size = 16
        Cells(i, 14).VerticalAlignment = xlVAlignCenter
        Cells(i, 14).HorizontalAlignment = xlHAlignCenter
        Cells(i, 14).Font.Size = 16
    Next i


End Sub
