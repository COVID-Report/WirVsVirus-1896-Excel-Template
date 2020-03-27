VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Change(ByVal Target As Range)
    'VBA liefert keine zuverlässige RowCount-Funktion daher wird vorher mit intRow die aktuelle Reihenzahl des nächsten Sheets bestimmt
    lrow = Worksheets(3).Cells(Rows.Count, 1).End(xlUp).Row
    lrow2 = Worksheets(4).Cells(Rows.Count, 1).End(xlUp).Row
    Dim intRow As Integer
    For intRow = 3 To lrow2
        If Worksheets(4).Cells(intRow, 1) = "" Then
            Exit For
        End If
    Next intRow
    
    'Wenn die 8. Spalte auf Positiv oder negativ gesetzt wird, wird die passende Userform aufgerufen und einige Werte übergeben
    Set KeyCellsTestStatus = Worksheets(3).Range(Worksheets(3).Cells(4, 8), Worksheets(3).Cells(lrow, 8))
    If Not Application.Intersect(KeyCellsTestStatus, Range(Target.Address)) Is Nothing And Target.Count = 1 Then
        If Target.Value = "Positiv - COVID-19 nachgewiesen" Then
           UserForm2.Label13.Caption = Target.Row 'CheatLabel
           UserForm2.Label7.Caption = Worksheets(3).Cells(Target.Row, 2).Value 'KrankenhausID
           UserForm2.Label8.Caption = Worksheets(3).Cells(Target.Row, 3).Value 'Vorname
           UserForm2.Label9.Caption = Worksheets(3).Cells(Target.Row, 4).Value 'Nachname
           UserForm2.Label10.Caption = Worksheets(3).Cells(Target.Row, 5).Value 'Geburtsdatum
           UserForm2.Show
        ElseIf Target.Value = "Negativ - COVID-19 nicht nachgewiesen" Then
           UserForm3.Label13.Caption = Target.Row 'CheatLabel
           UserForm3.Label7.Caption = Worksheets(3).Cells(Target.Row, 2).Value 'KrankenhausID
           UserForm3.Label8.Caption = Worksheets(3).Cells(Target.Row, 3).Value 'Vorname
           UserForm3.Label9.Caption = Worksheets(3).Cells(Target.Row, 4).Value 'Nachname
           UserForm3.Label10.Caption = Worksheets(3).Cells(Target.Row, 5).Value 'Geburtsdatum
           UserForm3.Show
        End If
    End If
    
End Sub

