Attribute VB_Name = "Module1"
Sub easyJet()
    
    Dim objHTTP As New MSXML2.XMLHTTP60
    Dim flightArray() As String
    
    URL = "http://www.easyjet.com/ejcms/cache15m/api/flights/search"
    
    URLwithParams = URL & "?AllDestinations=true&AllOrigins=true&AssumedPassengersPerBooking=1&AssumedSectorsPerBooking=1&CurrencyId=34&MaxResults=10000000&OriginIatas=BUD"
    
    objHTTP.Open "GET", URLwithParams, False
    objHTTP.send ("")
        
    flightArray = Split(objHTTP.ResponseText, "{")

    For i = 2 To UBound(flightArray)
        details = Split(flightArray(i), ",")
        Price = Split(details(0), ":")(1)
        airportTo = Mid(Split(details(2), ":")(1), 2, 3)
        DDate = Mid(Split(details(3), ":")(1), 2, 10)
        j = 2
        While Right(Cells(j, 1), 3) <> airportTo And Cells(j, 1) <> Empty
            j = j + 2
        Wend
        d = 2
        While Format(Cells(1, d).value, "yyyy-mm-dd") <> DDate And Cells(1, d) <> Empty
            d = d + 1
        Wend
        
        If Cells(j, 1) <> Empty And Cells(1, d) <> Empty Then
            Cells(j, d) = Price
        End If
    Next i
    
    URLwithParams = URL & "?AllDestinations=true&AllOrigins=true&AssumedPassengersPerBooking=1&AssumedSectorsPerBooking=1&CurrencyId=34&MaxResults=10000000&DestinationIatas=BUD"
    
    objHTTP.Open "GET", URLwithParams, False
    objHTTP.send ("")
    
    flightArray = Split(objHTTP.ResponseText, "{")

    For i = 2 To UBound(flightArray)
        details = Split(flightArray(i), ",")
        Price = Split(details(0), ":")(1)
        airportFrom = Mid(Split(details(1), ":")(1), 2, 3)
        DDate = Mid(Split(details(3), ":")(1), 2, 10)
        j = 2
        While Right(Cells(j, 1), 3) <> airportFrom And Cells(j, 1) <> Empty
            j = j + 2
        Wend
        d = 2
        While Format(Cells(1, d).value, "yyyy-mm-dd") <> DDate And Cells(1, d) <> Empty
            d = d + 1
        Wend
        
        If Cells(j, 1) <> Empty And Cells(1, d) <> Empty Then
            Cells(j + 1, d) = Price
        End If
    Next i
    
    Call Coloring
    
End Sub


Sub Coloring()
    
    j = 2
    While Cells(j, 1) <> Empty
        
        For t = 0 To 1
            
            i = 2
            While Cells(1, i) <> Empty
                
                Price = Cells(j + t, i).value
                If Price <> Empty Then
                    If Price <= 5000 Then
                        Cells(j + t, i).Interior.Color = RGB(94, 245, 87)
                    ElseIf Price <= 10000 Then
                        Cells(j + t, i).Interior.Color = RGB(129, 202, 74)
                    ElseIf Price <= 15000 Then
                        Cells(j + t, i).Interior.Color = RGB(172, 202, 74)
                    ElseIf Price <= 20000 Then
                        Cells(j + t, i).Interior.Color = RGB(202, 185, 74)
                    ElseIf Price <= 30000 Then
                        Cells(j + t, i).Interior.Color = RGB(219, 97, 97)
                    ElseIf Price > 30001 Then
                        Cells(j + t, i).Interior.Color = RGB(215, 18, 18)
                    End If
                Else
                    Cells(j + t, i) = "-"
                End If
                
                i = i + 1
            Wend
        Next t
        
        j = j + 2
    Wend
    
    
End Sub

Function MoneyToHUF(moneyCode As String) As Double

    Dim objHTTP As New MSXML2.XMLHTTP60

    'objHTTP.Open "GET", "http://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml", False
    
    objHTTP.Open "GET", "http://api.napiarfolyam.hu/?bank=otp", False
    objHTTP.send ("")
    
    from = InStr(objHTTP.ResponseText, moneyCode)
    If from <> 0 Then
        from = InStr(Mid(objHTTP.ResponseText, from, 100), "<eladas>") + from + 7
        fromTo = InStr(Mid(objHTTP.ResponseText, from, 100), "<")
        
        MoneyToHUF = CDbl(Replace(Mid(objHTTP.ResponseText, from, fromTo - 1), ".", ","))
    Else
        MoneyToHUF = 0
    End If
    
End Function

