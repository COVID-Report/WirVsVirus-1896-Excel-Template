VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Basisdaten eingeben"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9405.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
'Funktion fügt direkt in dem zweiten Sheet die Testpersonendaten ein, da bereits ein Abstrich erfolgt ist
lrow = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
Dim intRow As Integer
  For intRow = 3 To lrow
     If Cells(intRow, 1) = "" Then
        Exit For
     End If
  Next intRow

Worksheets(2).Cells(intRow, 1).Value = Format(Now, "dd-mm-yyyy hh:mm:ss") 'Zeitstempel
Worksheets(2).Cells(intRow, 2).Value = TextBox1.Value 'Krankenhaus-ID
Worksheets(2).Cells(intRow, 3).Value = TextBox2.Value 'Vorname
Worksheets(2).Cells(intRow, 4).Value = TextBox3.Value 'Nachname
Worksheets(2).Cells(intRow, 5).Value = TextBox4.Value 'Geburtsdatum
If OptionButton2.Value = True Then
    Worksheets(2).Cells(intRow, 6).Value = "SMS"
    ElseIf OptionButton3.Value = True Then
    Worksheets(2).Cells(intRow, 6).Value = "Mail"
    Else
    Worksheets(2).Cells(intRow, 6).Value = "Telefon"
End If
Worksheets(2).Cells(intRow, 7).Value = TextBox5.Value 'Telefonnummer
UserFormRefresh

'
'--------------------------------------
'ÜBERGABE VON HASH AND BACKEND
'--------------------------------------
'
End Sub

Private Sub CommandButton2_Click()
'Funktion fügt die Personendaten in das erste Sheet ein. Das erste Sheet kann als eine Wartezimmer verstanden werden mit allen Personen die registriert sind und auf den Abstrich warten
lrow = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
Dim intRow As Integer
  For intRow = 3 To lrow
     If Cells(intRow, 1) = "" Then
        Exit For
     End If
  Next intRow
Worksheets(1).Cells(intRow, 1).Value = Format(Now, "dd-mm-yyyy hh:mm:ss") 'Zeitstempel
Worksheets(1).Cells(intRow, 2).Value = TextBox1.Value 'Krankenhaus-ID
Worksheets(1).Cells(intRow, 3).Value = TextBox2.Value 'Vorname
Worksheets(1).Cells(intRow, 4).Value = TextBox3.Value 'Nachname
Worksheets(1).Cells(intRow, 5).Value = TextBox4.Value 'Geburtsdatum
If OptionButton2.Value = True Then
    Worksheets(1).Cells(intRow, 6).Value = "SMS"
    Else
    Worksheets(1).Cells(intRow, 6).Value = "Telefon"
End If
Worksheets(1).Cells(intRow, 7).Value = TextBox5.Value 'Telefonnummer
UserFormRefresh
'
'--------------------------------------
'ÜBERGABE VON HASH AND BACKEND
'--------------------------------------
'
End Sub

Private Sub CommandButton3_Click()
Unload UserForm1
End Sub

Private Sub OptionButton1_Change()
    If OptionButton1.Value = True Then
    Label6.Caption = "Telefonnummer"
    End If
End Sub

Private Sub OptionButton2_Change()
    If OptionButton2.Value = True Then
    Label6.Caption = "Handynummer"
    End If
End Sub

Private Sub OptionButton3_Change()
    If OptionButton3.Value = True Then
    Label6.Caption = "Email-Adresse"
    End If
End Sub

Private Sub TextBox4_AfterUpdate()
'Eine einfache Funktion, die es ermöglicht anstelle von 08.05.1992 einfach 08051992 einzugeben. Die Punkte werden danach automatisch eingefügt.
If InStr(TextBox4.Value, ".") = 0 And Len(TextBox4.Value) = 8 Then
TextBox4.Value = Left(TextBox4.Value, 2) + "." + Mid(TextBox4.Value, 3, 2) + "." + Right(TextBox4.Value, 4)
End If
End Sub

Sub UserFormRefresh()
    TextBox1.Value = ""
    TextBox2.Value = ""
    TextBox3.Value = ""
    TextBox4.Value = "dd.mm.yyyy"
    TextBox5.Value = ""
    TextBox1.SetFocus
    OptionButton1.Value = True
End Sub

Private Sub UserForm_Initialize()
    OptionButton1.Value = True
    TextBox4.Value = "dd.mm.yyyy"
End Sub


