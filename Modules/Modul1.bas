Attribute VB_Name = "Modul1"
Sub Schaltfläche1_Klicken()
UserForm1.Show
End Sub

Sub Schaltfläche3_Klicken()
    a = MsgBox("Wurden alle Abstriche gemacht?", vbYesNo)
    If a = vbYes Then
    lrow = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
    lrow2 = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
    Dim bool As Boolean
    Dim counter, intRow, intRow2 As Integer
    For intRow = 3 To lrow2
        If Worksheets(2).Cells(intRow, 1) = "" Then
            Exit For
        End If
    Next intRow
    For intRow2 = 3 To lrow2
        If Worksheets(1).Cells(intRow2, 1) = "" Then
            Exit For
        End If
    Next intRow2
    
    For counter = 4 To intRow2
        Worksheets(2).Cells(intRow + counter - 4, 1).Value = Worksheets(1).Cells(counter, 1).Value 'Zeitstempel
        Worksheets(2).Cells(intRow + counter - 4, 2).Value = Worksheets(1).Cells(counter, 2).Value 'KrankenhausID
        Worksheets(2).Cells(intRow + counter - 4, 3).Value = Worksheets(1).Cells(counter, 3).Value 'Vorname
        Worksheets(2).Cells(intRow + counter - 4, 4).Value = Worksheets(1).Cells(counter, 4).Value 'Nachname
        Worksheets(2).Cells(intRow + counter - 4, 5).Value = Worksheets(1).Cells(counter, 5).Value 'Geburtsdatum
        Worksheets(2).Cells(intRow + counter - 4, 6).Value = Worksheets(1).Cells(counter, 6).Value 'TEL/SMS
        Worksheets(2).Cells(intRow + counter - 4, 7).Value = Worksheets(1).Cells(counter, 7).Value 'Telefonnummer
    Next counter
    Worksheets(1).Range(Worksheets(1).Cells(4, 1), Worksheets(1).Cells(intRow2, 8)).Select
            Selection.Delete Shift:=xlUp
    End If
End Sub

Sub Schaltfläche6_Klicken()
lrow = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
Set KeyCellsTestStatus = Worksheets(1).Range(Worksheets(1).Cells(4, 1), Worksheets(1).Cells(lrow, 8))
If Not Application.Intersect(KeyCellsTestStatus, Range(Selection.Address)) Is Nothing Then
    a = MsgBox("Wurden alle Abstriche gemacht?", vbYesNo)
        If a = vbYes Then
        lrow = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
        lrow2 = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
        Dim bool As Boolean
        Dim counter, intRow, intRow2 As Integer
        For intRow = 3 To lrow2
            If Worksheets(2).Cells(intRow, 1) = "" Then
                Exit For
            End If
        Next intRow
        For intRow2 = 3 To lrow2
            If Worksheets(1).Cells(intRow2, 1) = "" Then
                Exit For
            End If
        Next intRow2
        
        For counter = 0 To Selection.Rows.Count
            Worksheets(2).Cells(intRow + counter, 1).Value = Worksheets(1).Cells(Selection.Row + counter, 1).Value 'Zeitstempel
            Worksheets(2).Cells(intRow + counter, 2).Value = Worksheets(1).Cells(Selection.Row + counter, 2).Value 'KrankenhausID
            Worksheets(2).Cells(intRow + counter, 3).Value = Worksheets(1).Cells(Selection.Row + counter, 3).Value 'Vorname
            Worksheets(2).Cells(intRow + counter, 4).Value = Worksheets(1).Cells(Selection.Row + counter, 4).Value 'Nachname
            Worksheets(2).Cells(intRow + counter, 5).Value = Worksheets(1).Cells(Selection.Row + counter, 5).Value 'Geburtsdatum
            Worksheets(2).Cells(intRow + counter, 6).Value = Worksheets(1).Cells(Selection.Row + counter, 6).Value 'TEL/SMS
            Worksheets(2).Cells(intRow + counter, 7).Value = Worksheets(1).Cells(Selection.Row + counter, 7).Value 'Telefonnummer
        Next counter
        Worksheets(1).Range(Worksheets(1).Cells(Selection.Row, 1), Worksheets(1).Cells(Selection.Row + Selection.Rows.Count - 1, 8)).Select
                Selection.Delete Shift:=xlUp
        End If
    End If
End Sub
Sub Schaltfläche4_Klicken()
    a = MsgBox("Sind alle Abstriche im Labor zur Untersuchung?", vbYesNo)
    If a = vbYes Then
    lrow = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
    lrow2 = Worksheets(3).Cells(Rows.Count, 1).End(xlUp).Row
    Dim bool As Boolean
    Dim counter, intRow, intRow2 As Integer
    For intRow = 3 To lrow2
        If Worksheets(3).Cells(intRow, 1) = "" Then
            Exit For
        End If
    Next intRow
    For intRow2 = 3 To lrow2
        If Worksheets(2).Cells(intRow2, 1) = "" Then
            Exit For
        End If
    Next intRow2
    
    For counter = 4 To intRow2
        Worksheets(3).Cells(intRow + counter - 4, 1).Value = Worksheets(2).Cells(counter, 1).Value 'Zeitstempel
        Worksheets(3).Cells(intRow + counter - 4, 2).Value = Worksheets(2).Cells(counter, 2).Value 'KrankenhausID
        Worksheets(3).Cells(intRow + counter - 4, 3).Value = Worksheets(2).Cells(counter, 3).Value 'Vorname
        Worksheets(3).Cells(intRow + counter - 4, 4).Value = Worksheets(2).Cells(counter, 4).Value 'Nachname
        Worksheets(3).Cells(intRow + counter - 4, 5).Value = Worksheets(2).Cells(counter, 5).Value 'Geburtsdatum
        Worksheets(3).Cells(intRow + counter - 4, 6).Value = Worksheets(2).Cells(counter, 6).Value 'TEL/SMS
        Worksheets(3).Cells(intRow + counter - 4, 7).Value = Worksheets(2).Cells(counter, 7).Value 'Telefonnummer
    Next counter
    Worksheets(2).Range(Worksheets(2).Cells(4, 1), Worksheets(2).Cells(intRow2 - 1, 8)).Select
            Selection.Delete Shift:=xlUp
    End If
End Sub
Sub Schaltfläche5_Klicken()
lrow = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
Set KeyCellsTestStatus = Worksheets(2).Range(Worksheets(2).Cells(4, 1), Worksheets(2).Cells(lrow, 8))
If Not Application.Intersect(KeyCellsTestStatus, Range(Selection.Address)) Is Nothing Then
    a = MsgBox("Sind die Abstriche im Labor zur Untersuchung?", vbYesNo)
        If a = vbYes Then
        lrow = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
        lrow2 = Worksheets(3).Cells(Rows.Count, 1).End(xlUp).Row
        Dim bool As Boolean
        Dim counter, intRow, intRow2 As Integer
        For intRow = 3 To lrow2
            If Worksheets(3).Cells(intRow, 1) = "" Then
                Exit For
            End If
        Next intRow
        For intRow2 = 3 To lrow2
            If Worksheets(2).Cells(intRow2, 1) = "" Then
                Exit For
            End If
        Next intRow2
        
        For counter = 0 To Selection.Rows.Count
            Worksheets(3).Cells(intRow + counter, 1).Value = Worksheets(2).Cells(Selection.Row + counter, 1).Value 'Zeitstempel
            Worksheets(3).Cells(intRow + counter, 2).Value = Worksheets(2).Cells(Selection.Row + counter, 2).Value 'KrankenhausID
            Worksheets(3).Cells(intRow + counter, 3).Value = Worksheets(2).Cells(Selection.Row + counter, 3).Value 'Vorname
            Worksheets(3).Cells(intRow + counter, 4).Value = Worksheets(2).Cells(Selection.Row + counter, 4).Value 'Nachname
            Worksheets(3).Cells(intRow + counter, 5).Value = Worksheets(2).Cells(Selection.Row + counter, 5).Value 'Geburtsdatum
            Worksheets(3).Cells(intRow + counter, 6).Value = Worksheets(2).Cells(Selection.Row + counter, 6).Value 'TEL/SMS
            Worksheets(3).Cells(intRow + counter, 7).Value = Worksheets(2).Cells(Selection.Row + counter, 7).Value 'Telefonnummer
        Next counter
        Worksheets(2).Range(Worksheets(2).Cells(Selection.Row, 1), Worksheets(2).Cells(Selection.Row + Selection.Rows.Count - 1, 8)).Select
                Selection.Delete Shift:=xlUp
        End If
    End If
End Sub

