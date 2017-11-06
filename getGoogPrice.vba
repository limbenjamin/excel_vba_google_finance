Function getGoogPrice(symbol As String) As Double
    Dim xmlhttp As Object
    Dim strURL As String
    Dim CompanyID As String
    Dim x As String
    Dim sSearch As String
     
    strURL = "https://finance.google.com/finance?q=" & symbol
    Set xmlhttp = CreateObject("msxml2.xmlhttp")
    With xmlhttp
        .Open "get", strURL, False
        .send
        x = .responsetext
    End With
    sSearch = "itemprop=""price"""
    priceMid = Mid(x, InStr(1, x, sSearch) + Len(sSearch) + 18)
    getGoogPrice = Left(priceMid, InStr(priceMid, """") - 1)

End Function

Private Sub Worksheet_Activate()
    Dim i As Integer

    For i = 2 To 10
      Cells(i, "D") = getGoogPrice(Cells(i, "E").Value)
    Next i
    'Cell E2:E10 should contain stock symbols e.g. SGX:Z74 for Singtel. Will populate cells D2:D10.
    
End Sub
