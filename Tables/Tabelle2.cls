VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Change(ByVal Target As Range)
    'VBA liefert keine zuverlässige RowCount-Funktion daher wird vorher mit intRow die aktuelle Reihenzahl des nächsten Sheets bestimmt
    lrow = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
    lrow2 = Worksheets(3).Cells(Rows.Count, 1).End(xlUp).Row
    Dim bool As Boolean
    Dim intRow As Integer
    For intRow = 3 To lrow2
        If Worksheets(3).Cells(intRow, 1) = "" Then
            Exit For
        End If
    Next intRow
    
    'Immer wenn die 8 Spalte auf "Im Labor zur Untersuchung" gesetzt wird, werden die Daten auf das nächste Sheet kopiert und aus dem alten Sheet entfernt
    Set KeyCellsTestStatus = Worksheets(2).Range(Worksheets(2).Cells(4, 8), Worksheets(2).Cells(lrow, 8))
    If Not Application.Intersect(KeyCellsTestStatus, Range(Target.Address)) Is Nothing And Target.Count = 1 Then
        If Target.Value = "Im Labor zur Untersuchung" Then
            'Hinzufügen der Daten zur Datenbank
            Tabelle1.posthash Worksheets(2).Cells(Target.Row, 5).Value, Worksheets(2).Cells(Target.Row, 2).Value, Worksheets(2).Cells(Target.Row, 4).Value, "PENDING", Worksheets(2).Cells(Target.Row, 7).Value
            'Übertragen der Daten auf das nächste Sheet
            Worksheets(3).Cells(intRow, 1).Value = Worksheets(2).Cells(Target.Row, 1).Value 'Aufgenommen am
            Worksheets(3).Cells(intRow, 2).Value = Worksheets(2).Cells(Target.Row, 2).Value 'Krankenhaus ID
            Worksheets(3).Cells(intRow, 3).Value = Worksheets(2).Cells(Target.Row, 3).Value 'Vorname
            Worksheets(3).Cells(intRow, 4).Value = Worksheets(2).Cells(Target.Row, 4).Value 'Nachname
            Worksheets(3).Cells(intRow, 5).Value = Worksheets(2).Cells(Target.Row, 5).Value 'Geburtsdatum
            Worksheets(3).Cells(intRow, 6).Value = Worksheets(2).Cells(Target.Row, 6).Value 'TEL/SMS
            Worksheets(3).Cells(intRow, 7).Value = Worksheets(2).Cells(Target.Row, 7).Value 'Telefonnummer
            'Entferne des Eintrages im "alten" Sheet
            Worksheets(2).Range(Worksheets(2).Cells(Target.Row, 1), Worksheets(2).Cells(Target.Row, 8)).Select
            Selection.Delete Shift:=xlUp
        End If
    End If
    
    '
    '--------------------------------------
    'ÜBERGABE VON HASH AND BACKEND
    '--------------------------------------
    '
End Sub
