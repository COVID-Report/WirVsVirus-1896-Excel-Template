VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
  
    'VBA liefert keine zuverlässige RowCount-Funktion daher wird vorher mit intRow die aktuelle Reihenzahl des nächsten Sheets bestimmt
    lrow = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
    lrow2 = Worksheets(2).Cells(Rows.Count, 1).End(xlUp).Row
    Dim bool As Boolean
    Dim intRow As Integer
    For intRow = 3 To lrow2
        If Worksheets(2).Cells(intRow, 1) = "" Then
            Exit For
        End If
    Next intRow
    
    'Immer wenn die 8 Spalte auf "Abstrich durchgeführt" gesetzt wird, werden die Daten auf das nächste Sheet kopiert und aus dem alten Sheet entfernt
    Set KeyCellsTestStatus = Worksheets(1).Range(Worksheets(1).Cells(4, 8), Worksheets(1).Cells(lrow, 8))
    If Not Application.Intersect(KeyCellsTestStatus, Range(Target.Address)) Is Nothing And Target.Count = 1 Then
        If Target.Value = "Abstrich durchgeführt" Then
            'Übertragen der Daten auf das nächste Sheet
            Worksheets(2).Cells(intRow, 1).Value = Worksheets(1).Cells(Target.Row, 1).Value 'Zeitstempel
            Worksheets(2).Cells(intRow, 2).Value = Worksheets(1).Cells(Target.Row, 2).Value 'KrankenhausID
            Worksheets(2).Cells(intRow, 3).Value = Worksheets(1).Cells(Target.Row, 3).Value 'Vorname
            Worksheets(2).Cells(intRow, 4).Value = Worksheets(1).Cells(Target.Row, 4).Value 'Nachname
            Worksheets(2).Cells(intRow, 5).Value = Worksheets(1).Cells(Target.Row, 5).Value 'Geburtsdatum
            Worksheets(2).Cells(intRow, 6).Value = Worksheets(1).Cells(Target.Row, 6).Value 'TEL/SMS
            Worksheets(2).Cells(intRow, 7).Value = Worksheets(1).Cells(Target.Row, 7).Value 'Telefonnummer
            'Entferne des Eintrages im "alten" Sheet
            Worksheets(1).Range(Worksheets(1).Cells(Target.Row, 1), Worksheets(1).Cells(Target.Row, 8)).Select
            Selection.Delete Shift:=xlUp
        End If
    End If

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
    strUser = "******"
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

Public Sub Testing()
    
    Dim newDate As String
    Dim hash As String
    datum = "08.02.1993"
    newDate = Format(datum, "YYYY-MM-DD")
    
    
    Dim objCryptoClass As clsSHA256
    Set objCryptoClass = New clsSHA256
    hash = objCryptoClass.SHA256("1234" + "Meyer" + newDate)
    
    Set objCryptoClass = Nothing
    Debug.Print "Hash:" + hash
    
End Sub
