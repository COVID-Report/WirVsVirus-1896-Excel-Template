VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   4515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8265.001
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
'Wenn das negative Ergebnis bestätigt wird, dann wird die Person direkt in das Sheet "abgeschlossene Fälle verschoben
lrow2 = Worksheets(5).Cells(Rows.Count, 1).End(xlUp).Row
    Dim intRow As Integer
    For intRow = 3 To lrow2
        If Worksheets(5).Cells(intRow, 1) = "" Then
            Exit For
        End If
    Next intRow
    'Hinzufügen der Daten zur Datenbank
    UserForm3.posthash Worksheets(3).Cells(Label13.Caption, 5).Value, Worksheets(3).Cells(Label13.Caption, 2).Value, Worksheets(3).Cells(Label13.Caption, 4).Value, "NEGATIVE", Worksheets(3).Cells(Label13.Caption, 7).Value
    Worksheets(5).Cells(intRow, 1).Value = Worksheets(3).Cells(Label13.Caption, 1).Value 'Angenommen am
    Worksheets(5).Cells(intRow, 2).Value = Worksheets(3).Cells(Label13.Caption, 2).Value 'Krankenhaus-ID
    Worksheets(5).Cells(intRow, 3).Value = Worksheets(3).Cells(Label13.Caption, 3).Value 'Vorname
    Worksheets(5).Cells(intRow, 4).Value = Worksheets(3).Cells(Label13.Caption, 4).Value 'Nachname
    Worksheets(5).Cells(intRow, 5).Value = Worksheets(3).Cells(Label13.Caption, 5).Value 'Geburtsdatum
    Worksheets(5).Cells(intRow, 6).Value = Worksheets(3).Cells(Label13.Caption, 6).Value 'TEL/SMS
    Worksheets(5).Cells(intRow, 7).Value = Worksheets(3).Cells(Label13.Caption, 7).Value 'Telefonnummer
    Worksheets(5).Cells(intRow, 8).Value = Worksheets(3).Cells(Label13.Caption, 8).Value 'Testergebnis
    Worksheets(5).Cells(intRow, 9).Value = Format(Now, "dd-mm-yyyy hh:mm:ss") 'Datum des Testergebnisses
    Worksheets(3).Range(Worksheets(3).Cells(Label13.Caption, 1), Worksheets(3).Cells(Label13.Caption, 8)).Select
    Selection.Delete Shift:=xlUp
    Unload UserForm3
    '
'--------------------------------------
'ÜBERGABE VON HASH AND BACKEND
'--------------------------------------
'
End Sub

Private Sub CommandButton2_Click()
'Wenn die Person eigentlich positiv ist, dann wird das falsche Formular geschlossen und das richtige Formular geöffnet und die Parameter übergeben
Unload UserForm3
Worksheets(3).Cells(Label13.Caption, 8).Value = "Positiv - COVID-19 nachgewiesen"
UserForm2.Label13.Caption = Label13.Caption
UserForm2.Label7.Caption = Worksheets(3).Cells(Label13.Caption, 2).Value 'Krankenhaus-ID
UserForm2.Label8.Caption = Worksheets(3).Cells(Label13.Caption, 3).Value 'Vorname
UserForm2.Label9.Caption = Worksheets(3).Cells(Label13.Caption, 4).Value 'Nachname
UserForm2.Label10.Caption = Worksheets(3).Cells(Label13.Caption, 5).Value 'Geburtsdatum
UserForm2.Show
End Sub

Private Sub CommandButton3_Click()
Worksheets(3).Cells(Label13.Caption, 8).Value = ""
Unload UserForm3
End Sub


Sub post(id As String, status As String, name As String, contact As String)
    Dim objRequest As Object
    Dim strUrl As String
    Dim blnAsync As Boolean
    Dim strResponse As String
    Dim strPost As String
    Set objRequest = CreateObject("MSXML2.XMLHTTP")
    Dim strUser As String
    Dim strPassword As String
    strUser = "****"
    strPassword = "*******"
    strUrl = "https://wirvsvirus-backend.azurewebsites.net/tests/" + id
    Debug.Print "Debug-Print id:" + id
    blnAsync = True
    strPost = "{" & Chr(10) & _
              """status""" & ": """ & status & """," & Chr(10) & _
              """name""" & ": """ & name & """," & Chr(10) & _
              """contact""" & ": """ & contact & """" & Chr(10) & _
              "}"
    Debug.Print "Debug-Print strPost:" + strPost
    With objRequest
        .Open "POST", strUrl, blnAsync
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Basic " & EncodeBase64(strUser & ":" & strPassword)
        .Send (strPost)
        'spin wheels whilst waiting for response
        While objRequest.readyState <> 4
            DoEvents
        Wend
        strResponse = .ResponseText
    End With
    Debug.Print "Debug-Print strResponse:" + strResponse
End Sub


Public Sub posthash(datum As String, krankenhausID As String, name As String, status As String, contact As String)
    
    Dim newDate As String
    Dim hash As String
    newDate = Format(datum, "YYYY-MM-DD")
    
    
    Dim objCryptoClass As clsSHA256
    Set objCryptoClass = New clsSHA256
    hash = objCryptoClass.SHA256(krankenhausID + name + newDate)
    
    Set objCryptoClass = Nothing
    post hash, status, name, contact
End Sub

Function EncodeBase64(text As String) As String
    Dim arrData() As Byte
    arrData = StrConv(text, vbFromUnicode)
    Dim objXML As Object
    Dim objNode As MSXML2.IXMLDOMElement
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = Application.Clean(objNode.text)
    Set objNode = Nothing
    Set objXML = Nothing
End Function


Private Sub Label7_Click()

End Sub
