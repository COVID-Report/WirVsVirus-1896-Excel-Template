VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

    'VBA liefert keine zuverlässige RowCount-Funktion daher wird vorher mit intRow die aktuelle Reihenzahl des nächsten Sheets bestimmt
    lrow = Worksheets(4).Cells(Rows.Count, 1).End(xlUp).Row
    lrow2 = Worksheets(5).Cells(Rows.Count, 1).End(xlUp).Row
    Dim intRow As Integer
    For intRow = 3 To lrow2
        If Worksheets(5).Cells(intRow, 1) = "" Then
            Exit For
        End If
    Next intRow
    
    'Wenn die 10. Spalte auf "Anruf getätigt" gesetzt wird, wird eine MsgBox mit JA oder Nein als Antwortmöglichkeiten aufgerufen.
    'Bei Ja werden die Daten in das Finale Sheet übertragen
    'Bei Nein wird der Status zurückgesetzt und muss erneut gesetzt werden
    
    Set KeyCellsTestStatus = Worksheets(4).Range(Worksheets(4).Cells(4, 10), Worksheets(4).Cells(lrow, 10))
    If Not Application.Intersect(KeyCellsTestStatus, Range(Target.Address)) Is Nothing And Target.Count = 1 Then
        If Target.Value = "Anruf getätigt" Then
            a = MsgBox("Wurde die Testperson angerufen?", vbYesNo)
            If a = vbYes Then
                Worksheets(5).Cells(intRow, 1).Value = Worksheets(4).Cells(Target.Row, 1).Value
                Worksheets(5).Cells(intRow, 2).Value = Worksheets(4).Cells(Target.Row, 2).Value
                Worksheets(5).Cells(intRow, 3).Value = Worksheets(4).Cells(Target.Row, 3).Value
                Worksheets(5).Cells(intRow, 4).Value = Worksheets(4).Cells(Target.Row, 4).Value
                Worksheets(5).Cells(intRow, 5).Value = Worksheets(4).Cells(Target.Row, 5).Value
                Worksheets(5).Cells(intRow, 6).Value = Worksheets(4).Cells(Target.Row, 6).Value
                Worksheets(5).Cells(intRow, 7).Value = Worksheets(4).Cells(Target.Row, 7).Value
                Worksheets(5).Cells(intRow, 8).Value = Worksheets(4).Cells(Target.Row, 8).Value
                Worksheets(5).Cells(intRow, 9).Value = Worksheets(4).Cells(Target.Row, 9).Value
                Worksheets(4).Range(Worksheets(4).Cells(Target.Row, 1), Worksheets(4).Cells(Target.Row, 10)).Select
                Selection.Delete Shift:=xlUp
            Else
                MsgBox Target.Row
                Worksheets(4).Cells(Target.Row, 10).Value = ""
            End If
        End If
    End If

End Sub

