Attribute VB_Name = "Mod_OptionStrat"
Option Explicit

' =================================================================================
'       MÓDULO CONSOLIDADO DE OPTIONSTRAT
'       Funciones para generar URLs de OptionStrat y convertirlas a órdenes
'
'       Consolida:
'       - OptStratURL_To_Order.bas (URLToOrder, funciones auxiliares)
'       - OptionStratURL_from_CSV.bas (GenerarURLsOptionStratConsolidado)
'       - Funciones de Options_Order_To_OptionStrart.bas (OptionStratURL, etc.)
' =================================================================================

' =================================================================================
' URL TO ORDER CONVERSION
' =================================================================================

' Preprocesses each leg token for proper quantity formatting
Private Function PreprocessLegToken(ByVal token As String, ByVal buildType As String, ByVal legIndex As Long) As String
    token = Trim(token)
    If InStr(token, "x") > 0 Then
        PreprocessLegToken = token
        Exit Function
    End If
    
    Select Case buildType
        Case "bull-put-spread", "bear-call-spread"
            If legIndex = 1 Then
                If Left(token, 2) <> "-." Then
                    If Left(token, 1) = "." Then
                        token = "-" & token
                    Else
                        token = "-." & token
                    End If
                End If
                PreprocessLegToken = token & "x1"
            Else
                If Left(token, 1) <> "." And Left(token, 2) <> "-." Then
                    token = "." & token
                End If
                PreprocessLegToken = token & "x1"
            End If
        Case Else
            If Left(token, 1) <> "." And Left(token, 2) <> "-." Then
                token = "." & token
            End If
            PreprocessLegToken = token & "x1"
    End Select
End Function

' Main function: Transforms OptionStrat URL to order format
Public Function URLToOrder(ByVal url As String) As String
    Dim buildType As String, baseUrl As String
    Dim lowerURL As String
    url = AddMissingX1(url)
    lowerURL = LCase(url)
    
    ' Determine strategy from URL
    If InStr(lowerURL, "build/bull-put-spread/") > 0 Then
        buildType = "bull-put-spread"
        baseUrl = "https://optionstrat.com/build/bull-put-spread/"
    ElseIf InStr(lowerURL, "build/bear-call-spread/") > 0 Then
        buildType = "bear-call-spread"
        baseUrl = "https://optionstrat.com/build/bear-call-spread/"
    ElseIf InStr(lowerURL, "build/long-call-butterfly/") > 0 Then
        buildType = "long-call-butterfly"
        baseUrl = "https://optionstrat.com/build/long-call-butterfly/"
    ElseIf InStr(lowerURL, "build/long-put-butterfly/") > 0 Then
        buildType = "long-put-butterfly"
        baseUrl = "https://optionstrat.com/build/long-put-butterfly/"
    ElseIf InStr(lowerURL, "build/call-broken-wing/") > 0 Then
        buildType = "call-broken-wing"
        baseUrl = "https://optionstrat.com/build/call-broken-wing/"
    ElseIf InStr(lowerURL, "build/put-broken-wing/") > 0 Then
        buildType = "put-broken-wing"
        baseUrl = "https://optionstrat.com/build/put-broken-wing/"
    ElseIf InStr(lowerURL, "build/iron-condor/") > 0 Then
        buildType = "iron-condor"
        baseUrl = "https://optionstrat.com/build/iron-condor/"
    Else
        buildType = "custom"
        baseUrl = "https://optionstrat.com/build/custom/"
    End If
    
    Dim urlTemp As String
    urlTemp = Replace(url, baseUrl, "")
    
    Dim slashPos As Long
    slashPos = InStr(urlTemp, "/")
    If slashPos = 0 Then
        URLToOrder = "Error: formato de URL inválido."
        Exit Function
    End If
    
    Dim ticker As String
    ticker = Left(urlTemp, slashPos - 1)
    Dim legsPart As String
    legsPart = Mid(urlTemp, slashPos + 1)
    
    Dim legs() As String
    legs = Split(legsPart, ",")
    Dim numLegs As Long
    numLegs = UBound(legs) - LBound(legs) + 1
    
    Dim i As Long
    For i = LBound(legs) To UBound(legs)
        legs(i) = PreprocessLegToken(legs(i), buildType, i - LBound(legs) + 1)
    Next i
    
    Dim legDate() As String, legStrike() As String, legQty() As Double, legOptLetter() As String, legCost() As Double
    ReDim legDate(1 To numLegs)
    ReDim legStrike(1 To numLegs)
    ReDim legQty(1 To numLegs)
    ReDim legOptLetter(1 To numLegs)
    ReDim legCost(1 To numLegs)
    
    For i = 1 To numLegs
        Dim token As String
        token = Trim(legs(i - 1))
        
        Dim costValue As Double
        costValue = 0
        Dim posAt As Long
        posAt = InStr(token, "@")
        If posAt > 0 Then
            Dim costStr As String
            costStr = Mid(token, posAt + 1)
            costStr = Replace(costStr, ",", ".")
            costValue = CDbl(costStr)
            token = Left(token, posAt - 1)
        End If
        legCost(i) = costValue
        
        If Left(token, 1) = "." Then
            token = Mid(token, 2)
        ElseIf Left(token, 2) = "-." Then
            token = Mid(token, 3)
        End If
        
        token = Mid(token, Len(ticker) + 1)
        
        If Len(token) < 6 Then
            URLToOrder = "Error: formato de dateCode inválido."
            Exit Function
        End If
        legDate(i) = Left(token, 6)
        token = Mid(token, 7)
        
        legOptLetter(i) = UCase(Left(token, 1))
        token = Mid(token, 2)
        
        Dim posX As Long
        posX = InStr(token, "x")
        If posX = 0 Then
            URLToOrder = "Error: formato de leg inválido (falta 'x')."
            Exit Function
        End If
        legStrike(i) = Left(token, posX - 1)
        Dim qtyStr As String
        qtyStr = Mid(token, posX + 1)
        On Error GoTo QtyError
        legQty(i) = CDbl(qtyStr)
        On Error GoTo 0
    Next i
    
    Dim netCost As Double
    netCost = 0
    For i = 1 To numLegs
        netCost = netCost + (legQty(i) * legCost(i))
    Next i
    
    Dim dateCode As String
    dateCode = legDate(1)
    Dim optLetter As String
    optLetter = UCase(legOptLetter(1))
    Dim optType As String
    If optLetter = "C" Then
        optType = "CALL"
    ElseIf optLetter = "P" Then
        optType = "PUT"
    Else
        optType = "UNKNOWN"
    End If
    
    Dim yy As String, mmStr As String, dd As String
    yy = Left(dateCode, 2)
    mmStr = Mid(dateCode, 3, 2)
    dd = Right(dateCode, 2)
    Dim monthName As String
    monthName = NumberToMonthAbbrev(CInt(mmStr))
    
    Dim weekStr As String
    weekStr = AddWeeklysIfNeeded(CInt(dd), CInt(mmStr), CInt(yy))
    If weekStr <> "" Then weekStr = weekStr & " "
    
    Dim orderType As String, baseQty As Long, strategy As String, strikeStrFinal As String, ratioStr As String
    
    Select Case numLegs
        Case 1 ' SIMPLE
            strategy = "SIMPLE"
            baseQty = Abs(legQty(1))
            strikeStrFinal = legStrike(1)
            If legQty(1) > 0 Then
                orderType = "BUY"
            Else
                orderType = "SELL"
            End If
            URLToOrder = orderType & " " & IIf(orderType = "SELL", "-" & baseQty, baseQty) & " " & _
                         ticker & " 100 " & weekStr & dd & " " & monthName & " " & yy & " " & _
                         strikeStrFinal & " " & optType & " @" & FormatCostValue(netCost) & " LMT"
                         
        Case 2 ' VERTICAL or BACKRATIO
            If Abs(legQty(1)) = Abs(legQty(2)) Then
                strategy = "VERTICAL"
                baseQty = Abs(legQty(1))
                strikeStrFinal = legStrike(1) & "/" & legStrike(2)
                If legQty(1) > 0 Then
                    orderType = "BUY"
                Else
                    orderType = "SELL"
                End If
                URLToOrder = orderType & " " & IIf(orderType = "SELL", "-" & baseQty, baseQty) & " " & _
                             strategy & " " & ticker & " 100 " & weekStr & dd & " " & monthName & " " & yy & " " & _
                             strikeStrFinal & " " & optType & " @" & FormatCostValue(netCost) & " LMT"
            Else
                strategy = "BACKRATIO"
                If legQty(1) > 0 Then
                    orderType = "SELL"
                Else
                    orderType = "BUY"
                End If
                baseQty = Abs(legQty(1))
                Dim sMult As Long, lMult As Long
                sMult = Abs(legQty(1)) \ baseQty
                lMult = Abs(legQty(2)) \ baseQty
                ratioStr = sMult & "/" & lMult
                strikeStrFinal = legStrike(1) & "/" & legStrike(2)
                URLToOrder = orderType & " " & IIf(orderType = "SELL", "-" & baseQty, baseQty) & " " & _
                             ratioStr & " " & strategy & " " & ticker & " 100 " & weekStr & dd & " " & monthName & " " & yy & " " & _
                             strikeStrFinal & " " & optType & " @" & FormatCostValue(netCost) & " LMT"
            End If
            
        Case 3 ' Butterfly / ~Butterfly
            Dim strikes(1 To 3) As Double, qtys(1 To 3) As Double
            Dim j As Long, temp As Variant
            For j = 1 To 3
                strikes(j) = CDbl(legStrike(j))
                qtys(j) = legQty(j)
            Next j
            
            Dim i2 As Long
            For i2 = 1 To 2
                For j = i2 + 1 To 3
                    If UCase(legOptLetter(1)) = "C" Then
                        If strikes(i2) > strikes(j) Then
                            temp = strikes(i2): strikes(i2) = strikes(j): strikes(j) = temp
                            temp = qtys(i2): qtys(i2) = qtys(j): qtys(j) = temp
                        End If
                    Else
                        If strikes(i2) < strikes(j) Then
                            temp = strikes(i2): strikes(i2) = strikes(j): strikes(j) = temp
                            temp = qtys(i2): qtys(i2) = qtys(j): qtys(j) = temp
                        End If
                    End If
                Next j
            Next i2
            
            strikeStrFinal = CStr(strikes(1)) & "/" & CStr(strikes(2)) & "/" & CStr(strikes(3))
            
            Dim gcdVal As Long, r1 As Long, r2 As Long, r3 As Long
            gcdVal = GCD3(Abs(qtys(1)), Abs(qtys(2)), Abs(qtys(3)))
            r1 = Abs(qtys(1)) / gcdVal
            r2 = Abs(qtys(2)) / gcdVal
            r3 = Abs(qtys(3)) / gcdVal
            
            If r2 = 3 And ((r1 = 1 And r3 = 2) Or (r1 = 2 And r3 = 1)) Then
                strategy = "~BUTTERFLY"
                If qtys(2) < 0 Then orderType = "BUY" Else orderType = "SELL"
                baseQty = gcdVal
                ratioStr = r1 & "/" & r2 & "/" & r3
                
                URLToOrder = orderType & " " & IIf(orderType = "BUY", "+", "-") & baseQty & " " & _
                             ratioStr & " " & strategy & " " & ticker & " 100 " & weekStr & dd & " " & monthName & " " & yy & " " & _
                             strikeStrFinal & " " & optType & " @" & FormatCostValue(netCost) & " LMT"
            
            ElseIf r1 = 1 And r2 = 2 And r3 = 1 Then
                strategy = "BUTTERFLY"
                If qtys(2) < 0 Then orderType = "BUY" Else orderType = "SELL"
                baseQty = gcdVal
                
                URLToOrder = orderType & " " & IIf(orderType = "BUY", "+", "-") & baseQty & " " & _
                             strategy & " " & ticker & " 100 " & weekStr & dd & " " & monthName & " " & yy & " " & _
                             strikeStrFinal & " " & optType & " @" & FormatCostValue(netCost) & " LMT"
            Else
                URLToOrder = "Error: Ratio de 3 piernas no reconocido (" & r1 & "/" & r2 & "/" & r3 & ")."
            End If
        
        Case 4 ' IRON CONDOR
            Dim putCount As Integer, callCount As Integer
            For j = 1 To 4
                If legOptLetter(j) = "P" Then putCount = putCount + 1
                If legOptLetter(j) = "C" Then callCount = callCount + 1
            Next j

            If Not (putCount = 2 And callCount = 2) Then
                URLToOrder = "Error: Estrategia de 4 patas no es un Iron Condor."
                Exit Function
            End If

            Dim putStrikes(1 To 2) As Double, callStrikes(1 To 2) As Double
            Dim pIdx As Integer: pIdx = 1
            Dim cIdx As Integer: cIdx = 1
            For j = 1 To 4
                If legOptLetter(j) = "P" Then
                    putStrikes(pIdx) = CDbl(legStrike(j))
                    pIdx = pIdx + 1
                Else
                    callStrikes(cIdx) = CDbl(legStrike(j))
                    cIdx = cIdx + 1
                End If
            Next j
            
            Dim lowPut As Double, highPut As Double, lowCall As Double, highCall As Double
            If putStrikes(1) < putStrikes(2) Then lowPut = putStrikes(1): highPut = putStrikes(2) Else lowPut = putStrikes(2): highPut = putStrikes(1)
            If callStrikes(1) < callStrikes(2) Then lowCall = callStrikes(1): highCall = callStrikes(2) Else lowCall = callStrikes(2): highCall = callStrikes(1)
            
            strikeStrFinal = CStr(lowPut) & "/" & CStr(highPut) & "/" & CStr(lowCall) & "/" & CStr(highCall)
            
            strategy = "IRON CONDOR"
            baseQty = Abs(legQty(1))
            
            If netCost < 0 Then
                orderType = "SELL"
            Else
                orderType = "BUY"
            End If
            
            URLToOrder = orderType & " " & IIf(orderType = "SELL", "-" & baseQty, baseQty) & " " & _
                         strategy & " " & ticker & " 100 " & weekStr & dd & " " & monthName & " " & yy & " " & _
                         strikeStrFinal & " @" & FormatCostValue(netCost) & " LMT"

        Case Else
            URLToOrder = "Error: número de piernas no soportado."
            Exit Function
    End Select
    Exit Function
QtyError:
    URLToOrder = "Error: cantidad inválida en una de las piernas."
End Function

' Helper function for weeklys detection
Private Function AddWeeklysIfNeeded(dayVal As Integer, monthVal As Integer, yy As Integer) As String
    Dim yearVal As Integer
    yearVal = 2000 + yy
    Dim expDate As Date
    expDate = DateSerial(yearVal, monthVal, dayVal)
    
    Dim firstDay As Date, firstFriday As Date, offset As Integer
    firstDay = DateSerial(yearVal, monthVal, 1)
    offset = (7 - Weekday(firstDay, vbFriday) + 1) Mod 7
    firstFriday = firstDay + offset
    Dim thirdFriday As Date
    thirdFriday = firstFriday + 14
    If expDate = thirdFriday Then
        AddWeeklysIfNeeded = ""
    Else
        AddWeeklysIfNeeded = "(Weeklys)"
    End If
End Function

' Format cost value for URL
Private Function FormatCostValue(costVal As Double) As String
    Dim sCost As String
    sCost = Format(costVal, "0.00#####")
    If Abs(costVal) < 1 Then
        If Left(sCost, 2) = "-0" Then
            sCost = "-" & Mid(sCost, 3)
        ElseIf Left(sCost, 1) = "0" Then
            sCost = Mid(sCost, 2)
        End If
    End If
    FormatCostValue = sCost
End Function

' Add missing x1 multipliers to URL
Public Function AddMissingX1(originalUrl As String) As String
    Dim baseUrl As String
    Dim pathPart As String
    Dim segments() As String
    Dim modifiedSegments() As String
    Dim i As Long
    Dim posSlash3 As Long
    Dim posAt As Long
    Dim checkPos As Long
    Dim charCheck As String
    Dim charBeforeMulti As String
    Dim currentSegment As String
    Dim modifiedPath As String
    Dim multiplierExists As Boolean

    posSlash3 = 0
    Dim slashCount As Integer: slashCount = 0
    Dim k As Long
    For k = 1 To Len(originalUrl)
        If Mid(originalUrl, k, 1) = "/" Then
            slashCount = slashCount + 1
            If slashCount = 3 Then
                posSlash3 = k
                Exit For
            End If
        End If
    Next k
    
    If posSlash3 = 0 Then
        AddMissingX1 = originalUrl
        Exit Function
    End If
    
    baseUrl = Left(originalUrl, posSlash3)
    pathPart = Mid(originalUrl, posSlash3 + 1)

    segments = Split(pathPart, ",")
    ReDim modifiedSegments(LBound(segments) To UBound(segments))

    For i = LBound(segments) To UBound(segments)
        currentSegment = Trim(segments(i))
        posAt = InStr(1, currentSegment, "@")
        multiplierExists = False

        If posAt > 1 Then
            checkPos = posAt - 1
            
            Do While checkPos >= 1
                charCheck = Mid(currentSegment, checkPos, 1)
                If IsNumeric(charCheck) Or charCheck = "." Then
                    checkPos = checkPos - 1
                Else
                    Exit Do
                End If
            Loop
            
            If checkPos >= 1 Then
                charCheck = Mid(currentSegment, checkPos, 1)

                If charCheck = "x" Then
                    If checkPos >= 2 Then
                        charBeforeMulti = Mid(currentSegment, checkPos - 1, 1)
                        If IsNumeric(charBeforeMulti) Or charBeforeMulti = "-" Then
                            multiplierExists = True
                        End If
                    End If
                    
                ElseIf charCheck = "-" Then
                     If checkPos >= 2 Then
                        charBeforeMulti = Mid(currentSegment, checkPos - 1, 1)
                        If charBeforeMulti = "x" Then
                            multiplierExists = True
                        End If
                    End If
                End If
            End If

            If Not multiplierExists Then
                modifiedSegments(i) = Left(currentSegment, posAt - 1) & "x1" & Mid(currentSegment, posAt)
            Else
                modifiedSegments(i) = currentSegment
            End If
            
        Else
            modifiedSegments(i) = currentSegment
        End If
    Next i

    modifiedPath = Join(modifiedSegments, ",")
    AddMissingX1 = baseUrl & modifiedPath
End Function

' =================================================================================
' GENERATE URLS FROM CSV
' =================================================================================

Sub GenerarURLsOptionStratConsolidado()
    Dim wkb As Workbook
    Dim wks As Worksheet
    Dim datos As Variant
    Dim fullCsvPath As String
    
    Dim startRow As Long, endRow As Long, i As Long
    Dim dict As Object
    
    Dim ticker As String
    Dim optionCode As String
    Dim qty As String
    Dim tradePrice As String
    Dim numQty As Double, numPrice As Double
    
    Dim legString As String
    Dim legCost As Double
    Dim key As Variant
    Dim finalURL As String
    Dim outputRow As Long

    Application.ScreenUpdating = False

    On Error GoTo FileErrorHandler
    
    fullCsvPath = GetCSVFilePath("AccountStatement.csv")
    
    If fullCsvPath <> "" Then
        Set wkb = Workbooks.Open(fullCsvPath)
    Else
        MsgBox "El archivo 'AccountStatement.csv' no se encontró.", vbCritical
        Exit Sub
    End If
    
    On Error GoTo 0

    Set wks = wkb.Sheets(1)
    datos = wks.UsedRange.value
    wkb.Close SaveChanges:=False

    startRow = 0
    endRow = 0
    For i = 1 To UBound(datos, 1)
        If LCase(Trim(datos(i, 1))) = "options" Then
            startRow = i + 1
        End If
        
        If startRow > 0 And LCase(Trim(datos(i, 2))) = "overall totals" Then
            endRow = i - 1
            Exit For
        End If
    Next i
    
    If startRow = 0 Or endRow = 0 Then
        MsgBox "No se pudo encontrar la tabla de 'Options'.", vbExclamation
        Exit Sub
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    For i = startRow + 1 To endRow
        If Not IsEmpty(datos(i, 1)) And Not IsEmpty(datos(i, 2)) Then
            ticker = Trim(CStr(datos(i, 1)))
            optionCode = Trim(CStr(datos(i, 2)))
            qty = Trim(CStr(datos(i, 6)))
            tradePrice = Trim(CStr(datos(i, 7)))
            
            If IsNumeric(qty) And IsNumeric(tradePrice) Then
                numQty = CDbl(qty)
                numPrice = CDbl(tradePrice)
            Else
                numQty = 0
                numPrice = 0
            End If
            
            legString = "." & optionCode & "x" & qty & "@" & tradePrice
            legCost = numQty * numPrice
            
            If dict.Exists(ticker) Then
                Dim dataArray As Variant
                dataArray = dict(ticker)
                dataArray(0) = dataArray(0) & "," & legString
                dataArray(1) = dataArray(1) + legCost
                dict(ticker) = dataArray
            Else
                dict.Add ticker, Array(legString, legCost)
            End If
        End If
    Next i
    
    Dim targetSheet As Worksheet
    Set targetSheet = ThisWorkbook.ActiveSheet
    
    outputRow = 10
    
    targetSheet.Range("B" & outputRow - 1 & ":D" & targetSheet.Rows.count).ClearContents
    targetSheet.Range("B" & outputRow - 1 & ":D" & outputRow - 1).value = Array("Ticker", "Costo", "URL Generada")

    For Each key In dict.Keys
        Dim finalData As Variant
        finalData = dict(key)
        
        Dim urlPart As String
        Dim totalCost As Double
        
        urlPart = finalData(0)
        totalCost = finalData(1) * 100
        
        finalURL = "https://optionstrat.com/build/custom/" & key & "/" & urlPart
        
        targetSheet.Cells(outputRow, "B").value = key
        targetSheet.Cells(outputRow, "C").value = totalCost
        targetSheet.Cells(outputRow, "D").value = finalURL
        
        outputRow = outputRow + 1
    Next key
    
    Dim statusMessage As String
    statusMessage = "Proceso completado: " & dict.count & " estrategias generadas." & vbCrLf & _
                    "Fecha de ejecuci�n: " & Format(Now, "dd/mm/yyyy hh:mm:ss") & vbCrLf & _
                    "Fecha del archivo CSV: " & Format(FileDateTime(fullCsvPath), "dd/mm/yyyy hh:mm:ss")

    With targetSheet.Range("D8")
        .ClearContents
        .value = statusMessage
        .WrapText = True
    End With

    Set dict = Nothing
    Set wks = Nothing
    Set wkb = Nothing
    Application.ScreenUpdating = True
    
    Exit Sub

FileErrorHandler:
    MsgBox "Error al abrir el archivo CSV." & vbCrLf & _
           "Verifique las rutas y que el archivo no esté abierto.", vbCritical
    Application.ScreenUpdating = True
End Sub

' =================================================================================
' ORDER TO OPTIONSTRAT URL
' =================================================================================

' Determines if operation is BUY or SELL
Public Function ParseIsBuy(ByVal lineText As String) As Boolean
    Dim upperLine As String: upperLine = UCase(lineText)
    If InStr(upperLine, "BUY") > 0 Or InStr(upperLine, "BOT") > 0 Then
        ParseIsBuy = True
    Else
        ParseIsBuy = False
    End If
End Function

' Reads net cost from line (e.g., "@-.71 LMT")
Public Function ParseNetCost(ByVal lineText As String) As Double
    Dim pos As Long: pos = InStr(1, lineText, "@")
    If pos = 0 Then
        ParseNetCost = 0
    Else
        Dim costStr As String
        costStr = Mid(lineText, pos + 1)
        If InStr(costStr, " ") > 0 Then costStr = Split(costStr, " ")(0)
        costStr = Replace(costStr, ",", ".")
        costStr = Trim(costStr)
        If Left(costStr, 2) = "-." Then
            costStr = "-0" & Mid(costStr, 2)
        ElseIf Left(costStr, 1) = "." Then
            costStr = "0" & costStr
        End If
        On Error Resume Next
        ParseNetCost = CDbl(costStr)
        On Error GoTo 0
    End If
End Function

' Checks if token is valid month abbreviation
Public Function MonthIsValid(ByVal token As String) As Boolean
    Dim m As String: m = UCase(Left(token, 3))
    Select Case m
        Case "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"
            MonthIsValid = True
        Case Else
            MonthIsValid = False
    End Select
End Function

' Builds OptionStrat URL from order description
Public Function OptionStratURL(ByVal InputLine As String) As String
    Dim baseUrl As String
    baseUrl = "https://optionstrat.com/build/custom/"
    
    Dim isBuy As Boolean: isBuy = ParseIsBuy(InputLine)
    
    Dim netCost As Double: netCost = ParseNetCost(InputLine)
    If Not isBuy Then netCost = -netCost
    
    Dim ticker As String, dateCode As String, QtyTrades As Long
    Dim allPositions() As Variant
    ParseSingleLineStrategy InputLine, ticker, dateCode, QtyTrades, allPositions
    
    Dim legPrices() As Double
    legPrices = DistributeCostAmongLegs(allPositions, netCost, QtyTrades)
    
    Dim finalLegs As String
    finalLegs = BuildLegsString(ticker, dateCode, allPositions, legPrices)
    
    OptionStratURL = baseUrl & ticker & "/" & finalLegs
End Function

' Parse single line strategy for OptionStratURL
Public Sub ParseSingleLineStrategy(ByVal lineText As String, _
                                    ByRef ticker As String, _
                                    ByRef dateCode As String, _
                                    ByRef QtyTrades As Long, _
                                    ByRef allPositions() As Variant)
    Dim tokens() As String
    tokens = Split(Trim(lineText), " ")
    
    Dim idx As Long: idx = 0
    Dim orderType As String: orderType = UCase(tokens(idx))
    idx = idx + 1
    
    Dim totalQty As Long: totalQty = Abs(CLng(tokens(idx)))
    QtyTrades = totalQty
    idx = idx + 1
    
    Dim ratioStr As String: ratioStr = ""
    Dim strategy As String: strategy = "SIMPLE"
    
    If idx <= UBound(tokens) Then
        Dim tokenVal As String: tokenVal = UCase(tokens(idx))
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
    
    Dim dayPart As String, monPart As String, yrPart As String
    dayPart = tokens(idx)
    monPart = tokens(idx + 1)
    yrPart = tokens(idx + 2)
    idx = idx + 3
    dateCode = MakeDateCode(dayPart, monPart, yrPart)
    
    Dim strikeInfo As String: strikeInfo = tokens(idx)
    idx = idx + 1
    Dim optType As String: optType = UCase(tokens(idx))
    idx = idx + 1
    
    Dim parts() As String, rTokens() As String
    Select Case strategy
        Case "SIMPLE"
            ReDim allPositions(1 To 1, 1 To 3)
            allPositions(1, 1) = CDbl(strikeInfo)
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
                allPositions(1, 1) = CDbl(parts(0))
                allPositions(1, 2) = optType
                allPositions(1, 3) = IIf(orderType = "BUY" Or orderType = "BOT", totalQty, -totalQty)
                
                allPositions(2, 1) = CDbl(parts(1))
                allPositions(2, 2) = optType
                allPositions(2, 3) = IIf(orderType = "BUY" Or orderType = "BOT", -totalQty, totalQty)
            Else
                ReDim allPositions(1 To 1, 1 To 3)
                allPositions(1, 1) = CDbl(strikeInfo)
                allPositions(1, 2) = optType
                allPositions(1, 3) = IIf(orderType = "BUY" Or orderType = "BOT", totalQty, -totalQty)
            End If
            
        Case "BUTTERFLY"
            parts = Split(strikeInfo, "/")
            Dim rLow As Long, rMid As Long, rHigh As Long
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
                allPositions(1, 1) = CDbl(parts(1))
                allPositions(1, 2) = optType
                allPositions(1, 3) = -rMid * totalQty
                
                allPositions(2, 1) = CDbl(parts(0))
                allPositions(2, 2) = optType
                allPositions(2, 3) = rLow * totalQty
                
                allPositions(3, 1) = CDbl(parts(2))
                allPositions(3, 2) = optType
                allPositions(3, 3) = rHigh * totalQty
            Else
                allPositions(1, 1) = CDbl(parts(0))
                allPositions(1, 2) = optType
                allPositions(1, 3) = -rLow * totalQty
                
                allPositions(2, 1) = CDbl(parts(1))
                allPositions(2, 2) = optType
                allPositions(2, 3) = rMid * totalQty
                
                allPositions(3, 1) = CDbl(parts(2))
                allPositions(3, 2) = optType
                allPositions(3, 3) = -rHigh * totalQty
            End If
            
        Case "BACKRATIO"
            parts = Split(strikeInfo, "/")
            Dim sMultiplier As Double, lMultiplier As Double
            sMultiplier = 1: lMultiplier = 1
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
                    allPositions(1, 1) = CDbl(parts(0))
                    allPositions(1, 2) = optType
                    allPositions(1, 3) = -sMultiplier * totalQty
                    
                    allPositions(2, 1) = CDbl(parts(1))
                    allPositions(2, 2) = optType
                    allPositions(2, 3) = lMultiplier * totalQty
                Else
                    allPositions(1, 1) = CDbl(parts(0))
                    allPositions(1, 2) = optType
                    allPositions(1, 3) = sMultiplier * totalQty
                    
                    allPositions(2, 1) = CDbl(parts(1))
                    allPositions(2, 2) = optType
                    allPositions(2, 3) = -lMultiplier * totalQty
                End If
            Else
                ReDim allPositions(1 To 1, 1 To 3)
                allPositions(1, 1) = CDbl(strikeInfo)
                allPositions(1, 2) = optType
                allPositions(1, 3) = IIf(orderType = "BUY" Or orderType = "BOT", totalQty, -totalQty)
            End If
            
        Case Else
            ReDim allPositions(1 To 1, 1 To 3)
            allPositions(1, 1) = CDbl(strikeInfo)
            allPositions(1, 2) = optType
            allPositions(1, 3) = IIf(orderType = "BUY" Or orderType = "BOT", totalQty, -totalQty)
    End Select
End Sub

' Distribute cost among legs for OptionStratURL
Public Function DistributeCostAmongLegs(allPos() As Variant, ByVal netCost As Double, ByVal totalQty As Long) As Double()
    Dim n As Long: n = UBound(allPos, 1)
    Dim prices() As Double: ReDim prices(1 To n)
    
    If netCost = 0 Then
        DistributeCostAmongLegs = prices
        Exit Function
    End If
    
    Dim i As Long, assigned As Boolean: assigned = False
    For i = 1 To n
        Dim qty As Double: qty = allPos(i, 3)
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

' Build legs string for URL
Public Function BuildLegsString(ByVal ticker As String, ByVal dateCode As String, _
                                 allPos() As Variant, prices() As Double) As String
    Dim n As Long: n = UBound(allPos, 1)
    Dim resultStr As String: resultStr = ""
    
    Dim i As Long
    For i = 1 To n
        Dim strike As Double
        strike = allPos(i, 1)
        
        Dim optType As String: optType = UCase(allPos(i, 2))
        Dim qty As Double: qty = allPos(i, 3)
        Dim priceVal As Double: priceVal = prices(i)
        Dim signQty As String: signQty = "x" & CStr(qty)
        
        Dim priceStr As String
        If priceVal = 0 Then
            priceStr = "@0"
        Else
            priceStr = "@" & Format(priceVal, "0.00#####")
        End If
        
        Dim legStr As String
        legStr = "." & ticker & dateCode & Left(optType, 1) & CStr(strike) & signQty & priceStr
        
        If resultStr = "" Then
            resultStr = legStr
        Else
            resultStr = resultStr & "," & legStr
        End If
    Next i
    
    BuildLegsString = resultStr
End Function

