VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub myMacro()
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        For j = 4 To 5
            If Cells(i, j) = "" Then
                Cells(i, j) = 0
            End If
        Next j
    Next i
            
End Sub
