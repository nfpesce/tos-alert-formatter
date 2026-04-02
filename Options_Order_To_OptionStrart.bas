Attribute VB_Name = "Options_Order_To_OptionStrart"
Option Explicit

' ------------------------------------------------------------------
' Función auxiliar que extrae el strike (numérico) de una cadena tipo "529 CALL"
' Soporta strikes con decimales usando punto o coma.
' ------------------------------------------------------------------
Public Function GetStrike(k As Variant) As Double
    Dim arr() As String
    
    On Error GoTo errHandler
    arr = Split(CStr(k), " ")
    GetStrike = ToDoubleInvariant(arr(0))
    Exit Function
    
errHandler:
    GetStrike = 0
End Function

' ------------------------------------------------------------------
' FUNCIÓN PRINCIPAL: SUMMARIZEOPTIONS
'   Recibe un rango con columnas: Cantidad, Ticker, Fecha de expiración,
'   y el detalle de la estrategia de opciones en la última columna.
'   Devuelve una cadena con las posiciones netas vigentes (no vencidas).
' ------------------------------------------------------------------
Public Function SummarizeOptions_old(FullRange As Range) As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, n As Long
    Dim q As Double
    Dim strat As String
    Dim resultStr2 As String
    Dim counter As Long
    Dim key As Variant
    Dim expDate As Variant
    
    Dim splitted() As String
    Dim ratioPart As String
    Dim ratioTokens() As String
    Dim r1 As Long, r2 As Long, r3 As Long
    Dim afterButterfly As String
    Dim tokens2() As String
    Dim subTokens2() As String
    
    Dim s As String
    Dim tokens() As String
    Dim subTokens() As String
    Dim strikesText As String
    Dim optType As String
    
    Dim itemCount As Long
    Dim sortedList() As Variant
    Dim idx As Long
    Dim i2 As Long, k2 As Long
    Dim tempKey As Variant, tempVal As Variant
    Dim currentQty As Double
    Dim currentKey As Variant
    Dim formattedQty As String
    
    n = FullRange.Rows.count
    
    For i = 1 To n
        
        If Not IsNumeric(FullRange.Cells(i, 1).value) Then GoTo NextRow
        q = FullRange.Cells(i, 1).value
        
        expDate = FullRange.Cells(i, 3).value
        
        If IsDate(expDate) Then
            If CDate(expDate) < Date Then GoTo NextRow
        End If
        
        If IsError(FullRange.Cells(i, FullRange.Columns.count).value) Then GoTo NextRow
        strat = Trim(FullRange.Cells(i, FullRange.Columns.count).value)
        If strat = "" Then GoTo NextRow
        
        If InStr(1, UCase(strat), "~BUTTERFLY") > 0 Then
            splitted = Split(strat, "~BUTTERFLY")
            
            ratioPart = Trim(splitted(0))
            ratioPart = Replace(ratioPart, "~", "")
            ratioPart = Trim(ratioPart)
            
            ratioTokens = Split(ratioPart, "/")
            r1 = 1: r2 = 2: r3 = 1
            
            If UBound(ratioTokens) = 2 Then
                r1 = CLng(ratioTokens(0))
                r2 = CLng(ratioTokens(1))
                r3 = CLng(ratioTokens(2))
            End If
            
            afterButterfly = Trim(splitted(1))
            tokens2 = Split(afterButterfly, " ")
            
            If UBound(tokens2) >= 1 Then
                strikesText = tokens2(0)
                optType = tokens2(1)
                
                subTokens2 = Split(strikesText, "/")
                If UBound(subTokens2) = 2 Then
                    key = subTokens2(0) & " " & optType
                    If dict.Exists(key) Then
                        dict(key) = dict(key) + q * r1
                    Else
                        dict.Add key, q * r1
                    End If
                    
                    key = subTokens2(1) & " " & optType
                    If dict.Exists(key) Then
                        dict(key) = dict(key) + q * (-r2)
                    Else
                        dict.Add key, q * (-r2)
                    End If
                    
                    key = subTokens2(2) & " " & optType
                    If dict.Exists(key) Then
                        dict(key) = dict(key) + q * r3
                    Else
                        dict.Add key, q * r3
                    End If
                End If
            End If
        
        ElseIf InStr(1, UCase(strat), "BUTTERFLY") > 0 Then
            s = Trim(Mid(strat, Len("BUTTERFLY ") + 1))
            tokens = Split(s, " ")
            
            If UBound(tokens) >= 1 Then
                strikesText = tokens(0)
                optType = tokens(1)
                subTokens = Split(strikesText, "/")
                
                If UBound(subTokens) >= 2 Then
                    key = subTokens(0) & " " & optType
                    If dict.Exists(key) Then
                        dict(key) = dict(key) + q
                    Else
                        dict.Add key, q
                    End If
                    
                    key = subTokens(1) & " " & optType
                    If dict.Exists(key) Then
                        dict(key) = dict(key) + q * (-2)
                    Else
                        dict.Add key, q * (-2)
                    End If
                    
                    key = subTokens(2) & " " & optType
                    If dict.Exists(key) Then
                        dict(key) = dict(key) + q
                    Else
                        dict.Add key, q
                    End If
                End If
            End If
            
        ElseIf InStr(1, UCase(strat), "VERTICAL") > 0 Then
            s = Trim(Mid(strat, Len("VERTICAL ") + 1))
            tokens = Split(s, " ")
            
            If UBound(tokens) >= 1 Then
                strikesText = tokens(0)
                optType = tokens(1)
                subTokens = Split(strikesText, "/")
                
                If UBound(subTokens) >= 1 Then
                    key = subTokens(0) & " " & optType
                    If dict.Exists(key) Then
                        dict(key) = dict(key) + q
                    Else
                        dict.Add key, q
                    End If
                    
                    key = subTokens(1) & " " & optType
                    If dict.Exists(key) Then
                        dict(key) = dict(key) + q * (-1)
                    Else
                        dict.Add key, q * (-1)
                    End If
                End If
            End If
            
        ElseIf InStr(1, UCase(strat), "BACKRATIO") > 0 Then
            Dim ratio As String
            Dim shortLegMultiplier As Double
            Dim longLegMultiplier As Double
            Dim s3 As String
            
            tokens = Split(strat, " ")
            If UBound(tokens) >= 3 Then
                ratio = tokens(0)
                subTokens = Split(ratio, "/")
                
                If UBound(subTokens) >= 1 Then
                    shortLegMultiplier = CDbl(subTokens(0))
                    longLegMultiplier = CDbl(subTokens(1))
                    
                    s3 = Trim(Mid(strat, InStr(1, UCase(strat), "BACKRATIO") + Len("BACKRATIO ")))
                    tokens = Split(s3, " ")
                    
                    If UBound(tokens) >= 1 Then
                        strikesText = tokens(0)
                        optType = tokens(1)
                        subTokens = Split(strikesText, "/")
                        
                        If UBound(subTokens) >= 1 Then
                            key = subTokens(0) & " " & optType
                            If dict.Exists(key) Then
                                dict(key) = dict(key) + q * (-shortLegMultiplier)
                            Else
                                dict.Add key, q * (-shortLegMultiplier)
                            End If
                            
                            key = subTokens(1) & " " & optType
                            If dict.Exists(key) Then
                                dict(key) = dict(key) + q * (longLegMultiplier)
                            Else
                                dict.Add key, q * (longLegMultiplier)
                            End If
                        End If
                    End If
                End If
            End If
            
        Else
            tokens = Split(strat, " ")
            If UBound(tokens) >= 1 Then
                key = tokens(0) & " " & tokens(1)
                If dict.Exists(key) Then
                    dict(key) = dict(key) + q
                Else
                    dict.Add key, q
                End If
            End If
        End If
        
NextRow:
    Next i
    
    itemCount = 0
    For Each key In dict.Keys
        If dict(key) <> 0 Then itemCount = itemCount + 1
    Next key
    
    If itemCount = 0 Then
        SummarizeOptions_old = "CLOSED"
        Exit Function
    End If
    
    ReDim sortedList(1 To itemCount, 1 To 2)
    
    idx = 1
    For Each key In dict.Keys
        If dict(key) <> 0 Then
            sortedList(idx, 1) = key
            sortedList(idx, 2) = dict(key)
            idx = idx + 1
        End If
    Next key
    
    For i2 = 1 To itemCount - 1
        For k2 = i2 + 1 To itemCount
            If GetStrike(sortedList(i2, 1)) > GetStrike(sortedList(k2, 1)) Then
                tempKey = sortedList(i2, 1)
                tempVal = sortedList(i2, 2)
                sortedList(i2, 1) = sortedList(k2, 1)
                sortedList(i2, 2) = sortedList(k2, 2)
                sortedList(k2, 1) = tempKey
                sortedList(k2, 2) = tempVal
            End If
        Next k2
    Next i2
    
    resultStr2 = ""
    counter = 0
    
    For i2 = 1 To itemCount
        currentQty = sortedList(i2, 2)
        currentKey = sortedList(i2, 1)
        counter = counter + 1
        
        If currentQty > 0 Then
            formattedQty = "+" & CStr(currentQty)
        Else
            formattedQty = CStr(currentQty)
        End If
        
        If resultStr2 = "" Then
            resultStr2 = formattedQty & "  " & currentKey
        Else
            resultStr2 = resultStr2 & " / " & formattedQty & "  " & currentKey
        End If
    Next i2
    
    If resultStr2 <> "" Then resultStr2 = resultStr2 & "."
    resultStr2 = Replace(Replace(resultStr2, "CALL", "C"), "PUT", "P")
    SummarizeOptions_old = resultStr2
End Function

' ------------------------------------------------------------------
' FUNCIÓN PRINCIPAL: NROCONTRATOS
'   Calcula el número total de contratos según la estrategia
'   (Se ha añadido lógica para ignorar contratos vencidos)
' ------------------------------------------------------------------
Public Function NroContratos_old(FullRange As Range) As Long
    Dim TotalContracts As Long
    Dim i As Long
    Dim expDate As Variant
    Dim isExpired As Boolean
    Dim qty As Long
    Dim StrategyText As String
    
    For i = 1 To FullRange.Rows.count
        
        expDate = FullRange.Cells(i, 3).value
        isExpired = False
        
        If IsDate(expDate) Then
            If CDate(expDate) < Date Then isExpired = True
        End If
        
        If Not isExpired Then
            qty = Abs(FullRange.Cells(i, 1).value)
            StrategyText = CStr(FullRange.Cells(i, FullRange.Columns.count).value)
            TotalContracts = TotalContracts + qty * ParseStrategy(StrategyText)
        End If
    Next i
    
    NroContratos_old = TotalContracts
End Function

' ------------------------------------------------------------------
' FUNCIÓN AUXILIAR: DETERMINA EL MULTIPLICADOR SEGÚN LA ESTRATEGIA
' ------------------------------------------------------------------
Public Function ParseStrategy(ByVal StrategyText As String) As Long
    Dim c As Long
    Dim upperText As String
    
    c = 0
    upperText = UCase(StrategyText)
    
    If InStr(upperText, "STOCK") > 0 Then
        c = 0
        
    ElseIf InStr(upperText, "BACKRATIO") > 0 Then
        Dim ratioPart As String
        ratioPart = Trim(Split(upperText, "BACKRATIO")(0))
        If ratioPart <> "" Then
            c = SumOfSplitted(ratioPart, "/")
        End If
        
    ElseIf InStr(upperText, "~") > 0 And InStr(upperText, "BUTTERFLY") > 0 Then
        Dim parts() As String
        Dim ratioPartButterfly As String
        
        parts = Split(upperText, "~")
        ratioPartButterfly = Trim(parts(0))
        
        If ratioPartButterfly <> "" Then
            c = SumOfSplitted(ratioPartButterfly, "/")
        Else
            c = 4
        End If
        
    ElseIf InStr(upperText, "~") > 0 And InStr(upperText, "CONDOR") > 0 Then
        Dim partsCondor() As String
        Dim ratioPartCondor As String
        
        partsCondor = Split(upperText, "~")
        ratioPartCondor = Trim(partsCondor(0))
        
        If ratioPartCondor <> "" Then
            c = SumOfSplitted(ratioPartCondor, "/")
        Else
            c = 4
        End If
        
    ElseIf InStr(upperText, "BUTTERFLY") > 0 Then
        c = 4
        
    ElseIf InStr(upperText, "IRON CONDOR") > 0 Then
        c = 4
        
    ElseIf InStr(upperText, "CONDOR") > 0 Then
        c = 4
        
    ElseIf InStr(upperText, "VERTICAL") > 0 Then
        c = 2
        
    ElseIf InStr(upperText, "PUT") > 0 Then
        c = 1
        
    ElseIf InStr(upperText, "CALL") > 0 Then
        c = 1
    End If
    
    ParseStrategy = c
End Function

' ------------------------------------------------------------------
' FUNCIÓN AUXILIAR: SUMA LOS VALORES DE UNA CADENA SEPARADA POR "/"
' ------------------------------------------------------------------
Public Function SumOfSplitted(ByVal textToSplit As String, ByVal delimiter As String) As Long
    Dim arr() As String
    Dim i As Long
    Dim sumVal As Long
    
    arr = Split(textToSplit, delimiter)
    
    For i = LBound(arr) To UBound(arr)
        If IsNumeric(arr(i)) Then
            sumVal = sumVal + CLng(arr(i))
        End If
    Next i
    
    SumOfSplitted = sumVal
End Function

' =======================================================================
' =============== FUNCIÓN: OptionStratURL ===============================
' Construye un URL para optionstrat.com a partir de una descripción
' =======================================================================
Public Function OptionStratURL_old(ByVal InputLine As String) As String
    Dim baseUrl As String
    Dim isBuy As Boolean
    Dim netCost As Double
    Dim ticker As String
    Dim dateCode As String
    Dim QtyTrades As Long
    Dim allPositions() As Variant
    Dim legPrices() As Double
    Dim finalLegs As String
    
    baseUrl = "https://optionstrat.com/build/custom/"
    
    isBuy = ParseIsBuy(InputLine)
    
    netCost = ParseNetCost(InputLine)
    If Not isBuy Then netCost = -netCost
    
    ParseSingleLineStrategy InputLine, ticker, dateCode, QtyTrades, allPositions
    
    legPrices = DistributeCostAmongLegs(allPositions, netCost, QtyTrades)
    finalLegs = BuildLegsString(ticker, dateCode, allPositions, legPrices)
    
    OptionStratURL_old = baseUrl & ticker & "/" & finalLegs
End Function

' ------------------------------------------------------------------
' Determina si la operación es BUY/BOT o SELL/SOLD
' ------------------------------------------------------------------
Public Function ParseIsBuy(ByVal lineText As String) As Boolean
    Dim upperLine As String
    upperLine = UCase(lineText)
    
    If InStr(upperLine, "BUY") > 0 Or InStr(upperLine, "BOT") > 0 Then
        ParseIsBuy = True
    Else
        ParseIsBuy = False
    End If
End Function

' ------------------------------------------------------------------
' Lee el costo neto desde la línea (ej. "@-.71 LMT")
' ------------------------------------------------------------------
Public Function ParseNetCost(ByVal lineText As String) As Double
    Dim pos As Long
    Dim costStr As String
    
    pos = InStr(1, lineText, "@")
    
    If pos = 0 Then
        ParseNetCost = 0
    Else
        costStr = Mid(lineText, pos + 1)
        
        If InStr(costStr, " ") > 0 Then costStr = Split(costStr, " ")(0)
        costStr = Trim(costStr)
        
        If Left(costStr, 2) = "-." Then
            costStr = "-0" & Mid(costStr, 2)
        ElseIf Left(costStr, 1) = "." Then
            costStr = "0" & costStr
        End If
        
        ParseNetCost = ToDoubleInvariant(costStr)
    End If
End Function

' ------------------------------------------------------------------
' Comprueba si un token es una abreviatura válida de mes (JAN, FEB, etc.)
' ------------------------------------------------------------------
Public Function MonthIsValid(ByVal token As String) As Boolean
    Dim m As String
    m = UCase(Left(token, 3))
    
    Select Case m
        Case "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"
            MonthIsValid = True
        Case Else
            MonthIsValid = False
    End Select
End Function

' ------------------------------------------------------------------
' PARSE DE LA LÍNEA PRINCIPAL PARA OptionStratURL
' ------------------------------------------------------------------
Public Sub ParseSingleLineStrategy(ByVal lineText As String, _
                                   ByRef ticker As String, _
                                   ByRef dateCode As String, _
                                   ByRef QtyTrades As Long, _
                                   ByRef allPositions() As Variant)
    Dim tokens() As String
    Dim idx As Long
    Dim orderType As String
    Dim totalQty As Long
    Dim ratioStr As String
    Dim strategy As String
    Dim tokenVal As String
    
    Dim dayPart As String
    Dim monPart As String
    Dim yrPart As String
    
    Dim strikeInfo As String
    Dim optType As String
    Dim parts() As String
    Dim rTokens() As String
    Dim rLow As Long, rMid As Long, rHigh As Long
    Dim sMultiplier As Double, lMultiplier As Double
    
    tokens = Split(Trim(lineText), " ")
    idx = 0
    
    orderType = UCase(tokens(idx))
    idx = idx + 1
    
    totalQty = Abs(CLng(tokens(idx)))
    QtyTrades = totalQty
    idx = idx + 1
    
    ratioStr = ""
    strategy = "SIMPLE"
    
    If idx <= UBound(tokens) Then
        tokenVal = UCase(tokens(idx))
        
        If InStr(tokenVal, "/") > 0 Then
            ratioStr = tokens(idx)
            idx = idx + 1
            
            If idx <= UBound(tokens) Then
                tokenVal = UCase(tokens(idx))
                If InStr(tokenVal, "BUTTERFLY") > 0 Then
                    strategy = "BUTTERFLY"
                    idx = idx + 1
                ElseIf InStr(tokenVal, "VERTICAL") > 0 Then
                    strategy = "VERTICAL"
                    idx = idx + 1
                ElseIf InStr(tokenVal, "BACKRATIO") > 0 Then
                    strategy = "BACKRATIO"
                    idx = idx + 1
                Else
                    strategy = "SIMPLE"
                End If
            End If
            
            ticker = tokens(idx)
            idx = idx + 1
            
        ElseIf InStr(tokenVal, "BUTTERFLY") > 0 Then
            strategy = "BUTTERFLY"
            idx = idx + 1
            ticker = tokens(idx)
            idx = idx + 1
            
        ElseIf InStr(tokenVal, "VERTICAL") > 0 Then
            strategy = "VERTICAL"
            idx = idx + 1
            ticker = tokens(idx)
            idx = idx + 1
            
        ElseIf InStr(tokenVal, "BACKRATIO") > 0 Then
            strategy = "BACKRATIO"
            idx = idx + 1
            ticker = tokens(idx)
            idx = idx + 1
            
        Else
            strategy = "SIMPLE"
            ticker = tokens(idx)
            idx = idx + 1
        End If
    End If
    
    Do While idx <= UBound(tokens)
        If IsNumeric(tokens(idx)) Then
            If (idx + 2 <= UBound(tokens)) And MonthIsValid(tokens(idx + 1)) Then Exit Do
        End If
        idx = idx + 1
    Loop
    
    dayPart = tokens(idx)
    monPart = tokens(idx + 1)
    yrPart = tokens(idx + 2)
    idx = idx + 3
    
    dateCode = MakeDateCode_old(dayPart, monPart, yrPart)
    
    strikeInfo = tokens(idx)
    idx = idx + 1
    optType = UCase(tokens(idx))
    idx = idx + 1
    
    Select Case strategy
        
        Case "SIMPLE"
            ReDim allPositions(1 To 1, 1 To 3)
            allPositions(1, 1) = ToDoubleInvariant(strikeInfo)
            allPositions(1, 2) = optType
            If orderType = "BUY" Or orderType = "BOT" Then
                allPositions(1, 3) = totalQty
            Else
                allPositions(1, 3) = -totalQty
            End If
            
        Case "VERTICAL"
            parts = Split(strikeInfo, "/")
            
            If UBound(parts) >= 1 Then
                ReDim allPositions(1 To 2, 1 To 3)
                allPositions(1, 1) = ToDoubleInvariant(parts(0))
                allPositions(1, 2) = optType
                allPositions(1, 3) = IIf(orderType = "BUY" Or orderType = "BOT", totalQty, -totalQty)
                
                allPositions(2, 1) = ToDoubleInvariant(parts(1))
                allPositions(2, 2) = optType
                allPositions(2, 3) = IIf(orderType = "BUY" Or orderType = "BOT", -totalQty, totalQty)
            Else
                ReDim allPositions(1 To 1, 1 To 3)
                allPositions(1, 1) = ToDoubleInvariant(strikeInfo)
                allPositions(1, 2) = optType
                If orderType = "BUY" Or orderType = "BOT" Then
                    allPositions(1, 3) = totalQty
                Else
                    allPositions(1, 3) = -totalQty
                End If
            End If
            
        Case "BUTTERFLY"
            parts = Split(strikeInfo, "/")
            rLow = 1: rMid = 2: rHigh = 1
            
            If ratioStr <> "" Then
                rTokens = Split(ratioStr, "/")
                If UBound(rTokens) = 2 Then
                    rLow = CLng(rTokens(0))
                    rMid = CLng(rTokens(1))
                    rHigh = CLng(rTokens(2))
                End If
            End If
            
            ReDim allPositions(1 To 3, 1 To 3)
            
            If orderType = "BUY" Or orderType = "BOT" Then
                allPositions(1, 1) = ToDoubleInvariant(parts(1))
                allPositions(1, 2) = optType
                allPositions(1, 3) = -rMid * totalQty
                
                allPositions(2, 1) = ToDoubleInvariant(parts(0))
                allPositions(2, 2) = optType
                allPositions(2, 3) = rLow * totalQty
                
                allPositions(3, 1) = ToDoubleInvariant(parts(2))
                allPositions(3, 2) = optType
                allPositions(3, 3) = rHigh * totalQty
            Else
                allPositions(1, 1) = ToDoubleInvariant(parts(0))
                allPositions(1, 2) = optType
                allPositions(1, 3) = -rLow * totalQty
                
                allPositions(2, 1) = ToDoubleInvariant(parts(1))
                allPositions(2, 2) = optType
                allPositions(2, 3) = rMid * totalQty
                
                allPositions(3, 1) = ToDoubleInvariant(parts(2))
                allPositions(3, 2) = optType
                allPositions(3, 3) = -rHigh * totalQty
            End If
            
        Case "BACKRATIO"
            parts = Split(strikeInfo, "/")
            sMultiplier = 1
            lMultiplier = 1
            
            If ratioStr <> "" Then
                rTokens = Split(ratioStr, "/")
                If UBound(rTokens) >= 1 Then
                    sMultiplier = CDbl(rTokens(0))
                    lMultiplier = CDbl(rTokens(1))
                End If
            End If
            
            If UBound(parts) >= 1 Then
                ReDim allPositions(1 To 2, 1 To 3)
                
                If orderType = "BUY" Or orderType = "BOT" Then
                    allPositions(1, 1) = ToDoubleInvariant(parts(0))
                    allPositions(1, 2) = optType
                    allPositions(1, 3) = -sMultiplier * totalQty
                    
                    allPositions(2, 1) = ToDoubleInvariant(parts(1))
                    allPositions(2, 2) = optType
                    allPositions(2, 3) = lMultiplier * totalQty
                Else
                    allPositions(1, 1) = ToDoubleInvariant(parts(0))
                    allPositions(1, 2) = optType
                    allPositions(1, 3) = sMultiplier * totalQty
                    
                    allPositions(2, 1) = ToDoubleInvariant(parts(1))
                    allPositions(2, 2) = optType
                    allPositions(2, 3) = -lMultiplier * totalQty
                End If
            Else
                ReDim allPositions(1 To 1, 1 To 3)
                allPositions(1, 1) = ToDoubleInvariant(strikeInfo)
                allPositions(1, 2) = optType
                If orderType = "BUY" Or orderType = "BOT" Then
                    allPositions(1, 3) = totalQty
                Else
                    allPositions(1, 3) = -totalQty
                End If
            End If
            
        Case Else
            ReDim allPositions(1 To 1, 1 To 3)
            allPositions(1, 1) = ToDoubleInvariant(strikeInfo)
            allPositions(1, 2) = optType
            If orderType = "BUY" Or orderType = "BOT" Then
                allPositions(1, 3) = totalQty
            Else
                allPositions(1, 3) = -totalQty
            End If
    End Select
End Sub

Public Function MakeDateCode_old(ByVal dayPart As String, ByVal monPart As String, ByVal yrPart As String) As String
    Dim dayVal As Long
    Dim monthNum As Long
    Dim yearVal As Long
    
    dayVal = CLng(dayPart)
    monthNum = MonthNumFromName(monPart)
    yearVal = CLng(yrPart)
    
    MakeDateCode_old = Format(yearVal, "00") & Format(monthNum, "00") & Format(dayVal, "00")
End Function

Public Function MonthNumFromName(ByVal mon As String) As Long
    Select Case UCase(Left(mon, 3))
        Case "JAN": MonthNumFromName = 1
        Case "FEB": MonthNumFromName = 2
        Case "MAR": MonthNumFromName = 3
        Case "APR": MonthNumFromName = 4
        Case "MAY": MonthNumFromName = 5
        Case "JUN": MonthNumFromName = 6
        Case "JUL": MonthNumFromName = 7
        Case "AUG": MonthNumFromName = 8
        Case "SEP": MonthNumFromName = 9
        Case "OCT": MonthNumFromName = 10
        Case "NOV": MonthNumFromName = 11
        Case "DEC": MonthNumFromName = 12
        Case Else: MonthNumFromName = 0
    End Select
End Function

' ------------------------------------------------------------------
' Distribuye el costo neto entre las piernas (para OptionStratURL).
' ------------------------------------------------------------------
Public Function DistributeCostAmongLegs(allPos() As Variant, ByVal netCost As Double, ByVal totalQty As Long) As Double()
    Dim n As Long
    Dim prices() As Double
    Dim i As Long
    Dim assigned As Boolean
    Dim qty As Double
    
    n = UBound(allPos, 1)
    ReDim prices(1 To n)
    
    If netCost = 0 Then
        DistributeCostAmongLegs = prices
        Exit Function
    End If
    
    assigned = False
    
    For i = 1 To n
        qty = allPos(i, 3)
        If Not assigned Then
            If (netCost > 0 And qty > 0) Or (netCost < 0 And qty < 0) Then
                prices(i) = Abs(netCost) / Abs(qty) * totalQty
                assigned = True
            End If
        Else
            prices(i) = 0
        End If
    Next i
    
    DistributeCostAmongLegs = prices
End Function

' ------------------------------------------------------------------
' Construye la cadena final de piernas para la URL (OptionStratURL).
' ------------------------------------------------------------------
Public Function BuildLegsString(ByVal ticker As String, ByVal dateCode As String, _
                                allPos() As Variant, prices() As Double) As String
    Dim n As Long
    Dim resultStr As String
    Dim i As Long
    Dim strike As Double
    Dim optType As String
    Dim qty As Double
    Dim priceVal As Double
    Dim signQty As String
    Dim priceStr As String
    Dim legStr As String
    Dim strikeStr As String
    
    n = UBound(allPos, 1)
    resultStr = ""
    
    For i = 1 To n
        strike = allPos(i, 1)
        optType = UCase(allPos(i, 2))
        qty = allPos(i, 3)
        priceVal = prices(i)
        signQty = "x" & CStr(qty)
        
        If priceVal = 0 Then
            priceStr = "@0"
        Else
            priceStr = "@" & Replace(Format(priceVal, "0.00#####"), ",", ".")
        End If
        
        strikeStr = DoubleToDotString(strike)
        legStr = "." & ticker & dateCode & Left(optType, 1) & strikeStr & signQty & priceStr
        
        If resultStr = "" Then
            resultStr = legStr
        Else
            resultStr = resultStr & "," & legStr
        End If
    Next i
    
    BuildLegsString = resultStr
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' =======================================================================
' =============== FUNCIÓN PRINCIPAL: OptionStratURL_FullTrade ===========
' Construye un URL para optionstrat.com a partir de un trade completo
' con múltiples posiciones.
'
' AJUSTES:
' - Soporta legs en formato nuevo: Qty + CALL/PUT + Strike  (ej: "-2 CALL 180")
'   y mantiene compatibilidad con formato viejo: Qty + Strike + CALL/PUT.
' - Soporta vencimientos múltiples por leg (fecha al final del leg):
'     "-4 PUT 380 27 FEB."   ó   "-4 PUT 380 27 FEB 26"
' - Si un leg trae "DD MMM" sin año, usa:
'     - año de fecha_expiracion si es parseable, o
'     - infiere año (año actual / próximo) si fecha_expiracion no es parseable.
' - Si un leg no trae fecha, usa fecha_expiracion (si es válida).
' - Excluye del resultado final todos los legs vencidos.
' - NUEVO: preserva correctamente strikes con decimales, por ejemplo 312.5
' =======================================================================
Public Function OptionStratURL_FullTrade(ByVal ticker As String, _
                                         ByVal fecha_expiracion As String, _
                                         ByVal string_posiciones As String, _
                                         ByVal costo_trade As Double) As String
    Dim baseUrl As String
    baseUrl = "https://optionstrat.com/build/custom/"
    
    ' -------------------------------------------------------------------
    ' 1) Parsear fecha_expiracion por defecto, si existe
    ' -------------------------------------------------------------------
    Dim dateCodeDefault As String
    Dim defaultYearToken As String
    Dim defaultExpirationDate As Date
    Dim hasDefaultExpirationDate As Boolean
    
    Dim f As String
    Dim defaultRawParts() As String
    Dim defaultParts As Collection
    Dim part As Variant
    Dim tok As String
    
    dateCodeDefault = ""
    defaultYearToken = ""
    hasDefaultExpirationDate = False
    
    f = Trim(fecha_expiracion)
    f = Replace(f, ".", "")
    f = Replace(f, ",", "")
    
    Set defaultParts = New Collection
    
    If f <> "" Then
        defaultRawParts = Split(f, " ")
        For Each part In defaultRawParts
            tok = Trim(CStr(part))
            If tok <> "" Then defaultParts.Add tok
        Next part
    End If
    
    If defaultParts.count >= 3 Then
        If TryBuildOSDate(CStr(defaultParts(1)), CStr(defaultParts(2)), CStr(defaultParts(3)), defaultExpirationDate) Then
            hasDefaultExpirationDate = True
            dateCodeDefault = DateToOSDateCode(defaultExpirationDate)
            defaultYearToken = CStr(Year(defaultExpirationDate))
        End If
    ElseIf defaultParts.count = 2 Then
        If TryInferOSDate(CStr(defaultParts(1)), CStr(defaultParts(2)), defaultExpirationDate) Then
            hasDefaultExpirationDate = True
            dateCodeDefault = DateToOSDateCode(defaultExpirationDate)
            defaultYearToken = CStr(Year(defaultExpirationDate))
        End If
    End If
    
    Set defaultParts = Nothing
    
    ' -------------------------------------------------------------------
    ' 2) Procesar el string_posiciones
    ' -------------------------------------------------------------------
    If Trim(string_posiciones) = "" Then
        OptionStratURL_FullTrade = "Error: string_posiciones está vacío."
        Exit Function
    End If
    
    Dim legDefinitionStrings() As String
    legDefinitionStrings = Split(Trim(string_posiciones), "/")
    
    Dim numLegs As Long
    numLegs = UBound(legDefinitionStrings) - LBound(legDefinitionStrings) + 1
    
    ' parsedLegs:
    '   (i,1) Qty (Long)
    '   (i,2) Strike (String)
    '   (i,3) OptType ("C" / "P")
    '   (i,4) DateCode ("YYMMDD")
    Dim parsedLegs() As Variant
    ReDim parsedLegs(1 To numLegs, 1 To 4)
    
    Dim activeLegCount As Long
    activeLegCount = 0
    
    Dim i As Long
    Dim currentLegString As String
    Dim rawLegParts() As String
    Dim cleanLegPartsColl As Collection
    
    Dim tCount As Long
    Dim legDateCode As String
    Dim legExpiryDate As Date
    Dim hasLegExpiryDate As Boolean
    Dim dateTokens As Long
    Dim lastTok As String, prevTok As String, prev2Tok As String
    Dim endIdx As Long
    Dim qtyStr As String
    Dim legQtyValue As Long
    Dim t2 As String, t3 As String
    Dim isT2Opt As Boolean, isT3Opt As Boolean
    Dim strikeTok As String, optTok As String
    Dim normalizedOptType As String
    
    For i = 1 To numLegs
        currentLegString = Trim(legDefinitionStrings(LBound(legDefinitionStrings) + i - 1))
        currentLegString = Replace(currentLegString, vbTab, " ")
        currentLegString = Trim(currentLegString)
        
        If currentLegString = "" Then
            OptionStratURL_FullTrade = "Error: Leg vacío en string_posiciones."
            Exit Function
        End If
        
        rawLegParts = Split(currentLegString, " ")
        Set cleanLegPartsColl = New Collection
        
        For Each part In rawLegParts
            If Trim(CStr(part)) <> "" Then
                tok = CleanTrailingPunctuation(CStr(part))
                If tok <> "" Then cleanLegPartsColl.Add tok
            End If
        Next part
        
        tCount = cleanLegPartsColl.count
        
        If tCount < 3 Then
            OptionStratURL_FullTrade = "Error: Formato de leg inválido (partes insuficientes) en: '" & currentLegString & "'"
            Exit Function
        End If
        
        legDateCode = ""
        hasLegExpiryDate = False
        dateTokens = 0
        
        lastTok = UCase(CStr(cleanLegPartsColl(tCount)))
        prevTok = ""
        prev2Tok = ""
        
        If tCount >= 2 Then prevTok = UCase(CStr(cleanLegPartsColl(tCount - 1)))
        If tCount >= 3 Then prev2Tok = UCase(CStr(cleanLegPartsColl(tCount - 2)))
        
        ' Caso A: ... DD MMM YY|YYYY
        If IsNumeric(lastTok) And (Len(lastTok) = 2 Or Len(lastTok) = 4) Then
            If MonthNumFromNameOS(prevTok) <> 0 And IsNumeric(prev2Tok) Then
                If TryBuildOSDate(prev2Tok, prevTok, lastTok, legExpiryDate) Then
                    legDateCode = DateToOSDateCode(legExpiryDate)
                    hasLegExpiryDate = True
                    dateTokens = 3
                End If
            End If
        End If
        
        ' Caso B: ... DD MMM (sin año)
        If legDateCode = "" Then
            If MonthNumFromNameOS(lastTok) <> 0 And IsNumeric(prevTok) Then
                If defaultYearToken <> "" Then
                    If TryBuildOSDate(prevTok, lastTok, defaultYearToken, legExpiryDate) Then
                        legDateCode = DateToOSDateCode(legExpiryDate)
                        hasLegExpiryDate = True
                        dateTokens = 2
                    End If
                Else
                    If TryInferOSDate(prevTok, lastTok, legExpiryDate) Then
                        legDateCode = DateToOSDateCode(legExpiryDate)
                        hasLegExpiryDate = True
                        dateTokens = 2
                    End If
                End If
            End If
        End If
        
        If legDateCode = "" Then
            If Not hasDefaultExpirationDate Then
                OptionStratURL_FullTrade = "Error: El leg no incluye fecha y fecha_expiracion no es válida: '" & fecha_expiracion & "'. Leg: '" & currentLegString & "'"
                Exit Function
            End If
            
            legExpiryDate = defaultExpirationDate
            legDateCode = dateCodeDefault
            hasLegExpiryDate = True
        End If
        
        ' Excluir legs vencidos
        If hasLegExpiryDate Then
            If legExpiryDate < Date Then GoTo NextLeg
        Else
            OptionStratURL_FullTrade = "Error: No se pudo determinar la fecha de vencimiento del leg: '" & currentLegString & "'"
            Exit Function
        End If
        
        endIdx = tCount - dateTokens
        
        If endIdx < 3 Then
            OptionStratURL_FullTrade = "Error: Formato de leg inválido (sin suficientes tokens luego de fecha) en: '" & currentLegString & "'"
            Exit Function
        End If
        
        qtyStr = CStr(cleanLegPartsColl(1))
        
        On Error Resume Next
        legQtyValue = CLng(qtyStr)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            OptionStratURL_FullTrade = "Error: Cantidad inválida '" & qtyStr & "' en: '" & currentLegString & "'"
            Exit Function
        End If
        On Error GoTo 0
        
        t2 = UCase(CStr(cleanLegPartsColl(2)))
        t3 = UCase(CStr(cleanLegPartsColl(3)))
        
        isT2Opt = (Left(t2, 1) = "C" Or Left(t2, 1) = "P")
        isT3Opt = (Left(t3, 1) = "C" Or Left(t3, 1) = "P")
        
        strikeTok = ""
        optTok = ""
        
        ' Nuevo: Qty + CALL/PUT + Strike
        If isT2Opt And IsNumericInvariant(CStr(cleanLegPartsColl(3))) Then
            optTok = t2
            strikeTok = NormalizeNumericToken(CStr(cleanLegPartsColl(3)))
        
        ' Viejo: Qty + Strike + CALL/PUT
        ElseIf IsNumericInvariant(CStr(cleanLegPartsColl(2))) And isT3Opt Then
            strikeTok = NormalizeNumericToken(CStr(cleanLegPartsColl(2)))
            optTok = t3
        
        Else
            OptionStratURL_FullTrade = "Error: No se pudo interpretar leg (se espera 'Qty CALL/PUT Strike' o 'Qty Strike CALL/PUT') en: '" & currentLegString & "'"
            Exit Function
        End If
        
        If Left(optTok, 1) = "C" Then
            normalizedOptType = "C"
        ElseIf Left(optTok, 1) = "P" Then
            normalizedOptType = "P"
        Else
            OptionStratURL_FullTrade = "Error: Tipo de opción inválido '" & optTok & "' en: '" & currentLegString & "'"
            Exit Function
        End If
        
        If Not IsNumericInvariant(strikeTok) Then
            OptionStratURL_FullTrade = "Error: Strike inválido '" & strikeTok & "' en: '" & currentLegString & "'"
            Exit Function
        End If
        
        activeLegCount = activeLegCount + 1
        parsedLegs(activeLegCount, 1) = legQtyValue
        parsedLegs(activeLegCount, 2) = strikeTok
        parsedLegs(activeLegCount, 3) = normalizedOptType
        parsedLegs(activeLegCount, 4) = legDateCode
        
NextLeg:
        Set cleanLegPartsColl = Nothing
    Next i
    
    If activeLegCount = 0 Then
        OptionStratURL_FullTrade = "Error: Todas las opciones están vencidas."
        Exit Function
    End If
    
    ' -------------------------------------------------------------------
    ' 3) Asignar costo_trade a un solo leg vigente
    ' -------------------------------------------------------------------
    Dim finalLegsString As String
    Dim costAssigned As Boolean
    Dim legPrice As Double
    Dim legQtyForCostCalc As Long
    Dim legStrike As String
    Dim legOptType As String
    Dim legPriceStr As String
    Dim legUrlComponent As String
    Dim legToAssignCost As Long
    Dim tempLegQty As Long
    
    finalLegsString = ""
    costAssigned = False
    legToAssignCost = -1
    
    If costo_trade <> 0 Then
        If costo_trade < 0 Then
            For i = 1 To activeLegCount
                tempLegQty = CLng(parsedLegs(i, 1))
                If tempLegQty > 0 Then
                    legToAssignCost = i
                    Exit For
                End If
            Next i
        Else
            For i = 1 To activeLegCount
                tempLegQty = CLng(parsedLegs(i, 1))
                If tempLegQty < 0 Then
                    legToAssignCost = i
                    Exit For
                End If
            Next i
        End If
        
        If legToAssignCost = -1 Then
            For i = 1 To activeLegCount
                If CLng(parsedLegs(i, 1)) <> 0 Then
                    legToAssignCost = i
                    Exit For
                End If
            Next i
        End If
    End If
    
    ' -------------------------------------------------------------------
    ' 4) Construir URL final solo con legs vigentes
    ' -------------------------------------------------------------------
    Dim legQtyOut As Long
    
    For i = 1 To activeLegCount
        legQtyOut = CLng(parsedLegs(i, 1))
        legStrike = CStr(parsedLegs(i, 2))
        legOptType = CStr(parsedLegs(i, 3))
        
        legPrice = 0
        
        If Not costAssigned And legQtyOut <> 0 And i = legToAssignCost Then
            legQtyForCostCalc = CLng(parsedLegs(legToAssignCost, 1))
            If legQtyForCostCalc <> 0 Then
                legPrice = Abs(costo_trade / legQtyForCostCalc / 100)
                costAssigned = True
            Else
                legPrice = 0
            End If
        End If
        
        If legPrice = 0 Then
            legPriceStr = "@0"
        Else
            legPriceStr = "@" & Replace(Format(legPrice, "0.0000"), ",", ".")
        End If
        
        legUrlComponent = "." & ticker & CStr(parsedLegs(i, 4)) & legOptType & legStrike & "x" & CStr(legQtyOut) & legPriceStr
        
        If finalLegsString = "" Then
            finalLegsString = legUrlComponent
        Else
            finalLegsString = finalLegsString & "," & legUrlComponent
        End If
    Next i
    
    If costo_trade <> 0 And Not costAssigned And legToAssignCost = -1 Then
        OptionStratURL_FullTrade = "Error: No se pudo asignar el costo. Todos los legs vigentes tienen cantidad 0 o no se encontró leg elegible."
        Exit Function
    End If
    
    OptionStratURL_FullTrade = baseUrl & ticker & "/" & finalLegsString
End Function

' ------------------------------------------------------------------
' Construye una fecha VBA real a partir de DD MMM YY|YYYY
' ------------------------------------------------------------------
Public Function TryBuildOSDate(ByVal dayPart As String, _
                               ByVal monPart As String, _
                               ByVal yrPart As String, _
                               ByRef outDate As Date) As Boolean
    On Error GoTo ErrorHandler
    
    Dim d As Long
    Dim m As Long
    Dim y As Long
    
    d = CLng(dayPart)
    m = MonthNumFromNameOS(monPart)
    
    If m = 0 Then
        TryBuildOSDate = False
        Exit Function
    End If
    
    y = CLng(yrPart)
    If Len(Trim(yrPart)) <= 2 Then
        y = 2000 + y
    End If
    
    outDate = DateSerial(y, m, d)
    TryBuildOSDate = True
    Exit Function

ErrorHandler:
    TryBuildOSDate = False
End Function

' ------------------------------------------------------------------
' Infiera la próxima ocurrencia de DD MMM a partir de hoy
' ------------------------------------------------------------------
Public Function TryInferOSDate(ByVal dayPart As String, _
                               ByVal monPart As String, _
                               ByRef outDate As Date) As Boolean
    On Error GoTo ErrorHandler
    
    Dim d As Long
    Dim m As Long
    Dim y As Long
    
    d = CLng(dayPart)
    m = MonthNumFromNameOS(monPart)
    
    If m = 0 Then
        TryInferOSDate = False
        Exit Function
    End If
    
    y = Year(Date)
    outDate = DateSerial(y, m, d)
    
    If outDate < Date Then
        outDate = DateSerial(y + 1, m, d)
    End If
    
    TryInferOSDate = True
    Exit Function

ErrorHandler:
    TryInferOSDate = False
End Function

' ------------------------------------------------------------------
' Convierte una fecha VBA a YYMMDD
' ------------------------------------------------------------------
Public Function DateToOSDateCode(ByVal dt As Date) As String
    DateToOSDateCode = Format(Year(dt) Mod 100, "00") & _
                       Format(Month(dt), "00") & _
                       Format(Day(dt), "00")
End Function

' ------------------------------------------------------------------
' Convierte Día Mes Año a formato YYMMDD para OptionStrat
' ------------------------------------------------------------------
Public Function MakeDateCodeOS(ByVal dayPart As String, ByVal monPart As String, ByVal yrPart As String) As String
    On Error GoTo ErrorHandler
    
    Dim dayVal As Long
    Dim monthNum As Long
    Dim yrFormat As String
    
    dayVal = CLng(dayPart)
    monthNum = MonthNumFromNameOS(monPart)
    
    If monthNum = 0 Then
        MakeDateCodeOS = "Error"
        Exit Function
    End If
    
    If Len(yrPart) = 4 Then
        yrFormat = Right(yrPart, 2)
    ElseIf Len(yrPart) = 2 Then
        yrFormat = Format(CLng(yrPart), "00")
    Else
        MakeDateCodeOS = "Error"
        Exit Function
    End If
    
    MakeDateCodeOS = yrFormat & Format(monthNum, "00") & Format(dayVal, "00")
    Exit Function

ErrorHandler:
    MakeDateCodeOS = "Error"
End Function

' ------------------------------------------------------------------
' Obtiene el número de mes a partir del nombre/abreviatura del mes
' ------------------------------------------------------------------
Public Function MonthNumFromNameOS(ByVal mon As String) As Long
    Select Case UCase(Left(mon, 3))
        Case "JAN", "ENE": MonthNumFromNameOS = 1
        Case "FEB": MonthNumFromNameOS = 2
        Case "MAR": MonthNumFromNameOS = 3
        Case "APR", "ABR": MonthNumFromNameOS = 4
        Case "MAY": MonthNumFromNameOS = 5
        Case "JUN": MonthNumFromNameOS = 6
        Case "JUL": MonthNumFromNameOS = 7
        Case "AUG", "AGO": MonthNumFromNameOS = 8
        Case "SEP": MonthNumFromNameOS = 9
        Case "OCT": MonthNumFromNameOS = 10
        Case "NOV": MonthNumFromNameOS = 11
        Case "DEC", "DIC": MonthNumFromNameOS = 12
        Case Else: MonthNumFromNameOS = 0
    End Select
End Function

' ------------------------------------------------------------------
' Limpia solamente puntuación al final del token.
' No toca puntos decimales internos como en 312.5
' ------------------------------------------------------------------
Public Function CleanTrailingPunctuation(ByVal token As String) As String
    Dim s As String
    Dim lastChar As String
    
    s = Trim(token)
    
    Do While Len(s) > 0
        lastChar = Right$(s, 1)
        If lastChar = "." Or lastChar = "," Or lastChar = ";" Or lastChar = ":" Then
            s = Left$(s, Len(s) - 1)
        Else
            Exit Do
        End If
    Loop
    
    CleanTrailingPunctuation = s
End Function

' ------------------------------------------------------------------
' Normaliza números para que usen punto decimal en forma interna.
' Ejemplos:
'   312,5  -> 312.5
'   312.5  -> 312.5
' ------------------------------------------------------------------
Public Function NormalizeNumericToken(ByVal token As String) As String
    Dim s As String
    
    s = Trim(token)
    s = CleanTrailingPunctuation(s)
    s = Replace(s, " ", "")
    
    If InStr(s, ",") > 0 And InStr(s, ".") > 0 Then
        If InStrRev(s, ".") > InStrRev(s, ",") Then
            s = Replace(s, ",", "")
        Else
            s = Replace(s, ".", "")
            s = Replace(s, ",", ".")
        End If
    ElseIf InStr(s, ",") > 0 Then
        s = Replace(s, ",", ".")
    End If
    
    NormalizeNumericToken = s
End Function

' ------------------------------------------------------------------
' Verifica si un token representa un número usando punto o coma decimal
' ------------------------------------------------------------------
Public Function IsNumericInvariant(ByVal token As String) As Boolean
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim dotCount As Long
    
    s = NormalizeNumericToken(token)
    
    If s = "" Then
        IsNumericInvariant = False
        Exit Function
    End If
    
    If Left$(s, 1) = "+" Or Left$(s, 1) = "-" Then
        s = Mid$(s, 2)
    End If
    
    If s = "" Then
        IsNumericInvariant = False
        Exit Function
    End If
    
    dotCount = 0
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "." Then
            dotCount = dotCount + 1
            If dotCount > 1 Then
                IsNumericInvariant = False
                Exit Function
            End If
        ElseIf ch < "0" Or ch > "9" Then
            IsNumericInvariant = False
            Exit Function
        End If
    Next i
    
    If s = "." Then
        IsNumericInvariant = False
        Exit Function
    End If
    
    IsNumericInvariant = True
End Function

' ------------------------------------------------------------------
' Convierte token numérico normalizado a Double
' Usa Val para interpretar siempre punto decimal.
' ------------------------------------------------------------------
Public Function ToDoubleInvariant(ByVal token As String) As Double
    Dim s As String
    
    s = NormalizeNumericToken(token)
    
    If s = "" Then
        ToDoubleInvariant = 0
    Else
        ToDoubleInvariant = Val(s)
    End If
End Function

' ------------------------------------------------------------------
' Convierte un Double a string con punto decimal para URL
' ------------------------------------------------------------------
Public Function DoubleToDotString(ByVal value As Double) As String
    Dim s As String
    
    s = Trim$(Str$(value))
    s = Replace(s, ",", ".")
    
    DoubleToDotString = s
End Function

