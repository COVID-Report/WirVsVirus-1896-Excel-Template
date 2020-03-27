VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Testergebnis bestätigen"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
'Wenn das positive Ergebnis bestätigt wird, wird die Testperson in das Sheet "Positive Ergebnisse" verschoben
lrow2 = Worksheets(4).Cells(Rows.Count, 1).End(xlUp).Row
    Dim intRow As Integer
    Dim geb, kid, nname, contact As String
    For intRow = 3 To lrow2
        If Worksheets(4).Cells(intRow, 1) = "" Then
            Exit For
        End If
    Next intRow
    'Hinzufügen der Daten zur Datenbank
    UserForm2.posthash Worksheets(3).Cells(Label13.Caption, 5).Value, Worksheets(3).Cells(Label13.Caption, 2).Value, Worksheets(3).Cells(Label13.Caption, 4).Value, "POSITIVE", Worksheets(3).Cells(Label13.Caption, 7).Value
Worksheets(4).Cells(intRow, 1).Value = Worksheets(3).Cells(Label13.Caption, 1).Value 'Krankenhaus-ID
Worksheets(4).Cells(intRow, 2).Value = Worksheets(3).Cells(Label13.Caption, 2).Value 'Krankenhaus-ID
Worksheets(4).Cells(intRow, 3).Value = Worksheets(3).Cells(Label13.Caption, 3).Value 'Vorname
Worksheets(4).Cells(intRow, 4).Value = Worksheets(3).Cells(Label13.Caption, 4).Value 'Nachname
Worksheets(4).Cells(intRow, 5).Value = Worksheets(3).Cells(Label13.Caption, 5).Value 'Geburtsdatum
Worksheets(4).Cells(intRow, 6).Value = Worksheets(3).Cells(Label13.Caption, 6).Value 'TEL/SMS
Worksheets(4).Cells(intRow, 7).Value = Worksheets(3).Cells(Label13.Caption, 7).Value 'Telefonnummer
Worksheets(4).Cells(intRow, 8).Value = Worksheets(3).Cells(Label13.Caption, 8).Value 'Testergebnis
Worksheets(4).Cells(intRow, 9).Value = Format(Now, "dd-mm-yyyy hh:mm:ss") 'Datum des Testergebnisses
Worksheets(3).Range(Worksheets(3).Cells(Label13.Caption, 1), Worksheets(3).Cells(Label13.Caption, 8)).Select
Selection.Delete Shift:=xlUp
Unload UserForm2
'
'--------------------------------------
'ÜBERGABE VON HASH AND BACKEND
'--------------------------------------
'
End Sub

Private Sub CommandButton2_Click()
'Wenn die Person nicht positiv sondern negativ ist und dies in dem Userformular auffällt, so wird diese Funktion aufgerufen.
'Die Paramater werden an das passende Formular übergeben und das Formular wird aufgerufen
Unload UserForm2
UserForm3.Label13.Caption = Label13.Caption
Worksheets(3).Cells(Label13.Caption, 8).Value = "Negativ - COVID-19 nicht nachgewiesen"
UserForm3.Label7.Caption = Worksheets(3).Cells(Label13.Caption, 2).Value 'Krankenhaus-ID
UserForm3.Label8.Caption = Worksheets(3).Cells(Label13.Caption, 3).Value 'Vornmae
UserForm3.Label9.Caption = Worksheets(3).Cells(Label13.Caption, 4).Value 'Nachname
UserForm3.Label10.Caption = Worksheets(3).Cells(Label13.Caption, 5).Value 'Geburtsdatum
UserForm3.Show
End Sub

Private Sub CommandButton3_Click()
Worksheets(3).Cells(Label13.Caption, 8).Value = ""
Unload UserForm2
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
    strUser = "*****"
    strPassword = "****"
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


