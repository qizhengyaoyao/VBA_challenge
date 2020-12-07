Sub Stock():


    Dim sht As Worksheet
    'Set sht = ActiveSheet
    'MsgBox (LastRow)
    For Each sht In Worksheets
    
        sht.Activate
        
        StatClHeader
        
        StatCl
        
        StatAlHeader
        
        StatAl
    
    Next sht

End Sub

Sub StatClHeader()
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Changed"
    Cells(1, 11).Value = "Percent Changed"
    Cells(1, 12).Value = "Total Stock Volume"

End Sub

Sub StatCl() 'sht As Worksheet)
    
    Dim LastRow As Long
    Dim sht As Worksheet
    Set sht = ActiveSheet
    LastRow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row

    Dim trow As Long
    Dim opprice As Double
    Dim clprice As Double
    Dim totvol As LongLong
    
    trow = 2
    opprice = Cells(2, 3)
    totvol = 0
    
    Columns("K").NumberFormat = "0.00%"
    
    For i = 2 To LastRow
        If Cells(i, 1) <> Cells(i + 1, 1) Then
            totvol = totvol + Cells(i, 7).Value
            clprice = Cells(i, 6).Value
            Cells(trow, 9).Value = Cells(i, 1).Value
            Cells(trow, 10).Value = clprice - opprice
            If opprice = 0 Then
                Cells(trow, 11).Value = 0
            Else
                Cells(trow, 11).Value = (clprice - opprice) / opprice
            End If
            
            Cells(trow, 12).Value = totvol
            
            trow = trow + 1
            opprice = Cells(i + 1, 3)
            totvol = 0
        Else
            totvol = totvol + Cells(i, 7).Value
        End If
    Next i
End Sub

Sub StatAlHeader()
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % increase"
    Cells(3, 14).Value = "Greatest % decrease"
    Cells(4, 14).Value = "Greatest total volume"
End Sub

Sub StatAl() 'sht As Worksheet)
    Dim LastRow As Long
    Dim sht As Worksheet
    Set sht = ActiveSheet
    LastRow = sht.Cells(sht.Rows.Count, "I").End(xlUp).Row
    
    Dim Tinc, Tdec, Ttot As String
    Dim Ninc, Ndec As Double
    Dim Ntot As LongLong
    Tinc = ""
    Tdec = ""
    Ttot = ""
    Ninc = 0
    Ndec = 0
    Ntot = 0
    
    
    For i = 2 To LastRow
        If Ninc < Cells(i, 11) Then
            Tinc = Cells(i, 9).Value
            Ninc = Cells(i, 11).Value
        End If
        
        If Ndec > Cells(i, 11) Then
            Tdec = Cells(i, 9).Value
            Ndec = Cells(i, 11).Value
        End If
        
        If Ntot < Cells(i, 12) Then
            Ttot = Cells(i, 9).Value
            Ntot = Cells(i, 12).Value
        End If
    Next i
    
    Range("P2:P3").NumberFormat = "0.00%"
    Cells(2, 15) = Tinc
    Cells(3, 15) = Tdec
    Cells(4, 15) = Ttot
    Cells(2, 16) = Ninc
    Cells(3, 16) = Ndec
    Cells(4, 16) = Ntot
    
End Sub

