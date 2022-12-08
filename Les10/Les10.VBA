Option Explicit
'This code is the solution to lesson 10
'Ferre Meuleman
'24/11/2022

Sub Oef10_1()
    Dim intA As Integer, intB As Integer
    intA = CInt(InputBox("Geef een integergetal in"))
    intB = intDubbel(intA)
    MsgBox ("Het resultaat is " & CStr(intB))
End Sub

Function intDubbel(intA) As Integer
    intDubbel = intA * 2
End Function


Sub Oef10_2()
    Dim intA As Integer, intB As Integer
    intA = intlees()
    intB = intDubbel(intA)
    MsgBox ("Het resultaat is " & CStr(intB))
End Sub

Function intlees() As String
    intlees = CInt(InputBox("Geef een integergetal in"))
End Function


Sub Oef10_3()
        Dim intA As Integer, intB As Integer
intA = intlees()
        intB = intDubbel(intA)
            Call drukAf(intB)
End Sub

Sub drukAf(intB)
    MsgBox ("Het resultaat is " & CStr(intB))
End Sub

Sub Oef10_4()
    Const MIN = 5
    Const MAX = 15
    Dim intA As Integer
    intA = CInt(InputBox("Geef getal"))
    If boolGetalTussen(intA, MIN, MAX) Then
        MsgBox ("OK")
    Else
        MsgBox ("NOK")
    End If
End Sub

Function boolGetalTussen(intA, MIN, MAX) As Boolean
    If intA <= MAX And intA >= MIN Then
        boolGetalTussen = True
    Else
        boolGetalTussen = False
    End If

End Function


Sub Oef10_5()
    Dim strIn1 As String
    Dim strIn2 As String
    Dim strIn3 As String
    Dim strOut As String
    strIn1 = strLeesGetalStringGrenzen("Geef minimum", 0, 100)
    strIn2 = strLeesGetalStringGrenzen("Geef maximum", 0, 100)
    strIn3 = strLeesGetalStringGrenzen("Geef Getal", CInt(strIn1), CInt(strIn2))
    strOut = "U gaf in " & vbNewLine
    strOut = strOut & "Minimum: " & strIn1 & vbNewLine
    strOut = strOut & "Maximum: " & strIn2 & vbNewLine
    strOut = strOut & "Getal: " & strIn3
    MsgBox (strOut)
End Sub

Function strLeesGetalStringGrenzen(strA, intA, intB) As String
    Dim intNumA As Integer
    
    Do
        intNumA = CInt(Inputopnemen(strA, intA, intB))

    Loop Until intA < intNumA And intNumA < intB

    strLeesGetalStringGrenzen = CStr(intNumA)
    

End Function

Function Inputopnemen(strA, intA, intB) As String
    Do
        Inputopnemen = InputBox(strA & " tussen " & CStr(intA) & " en " & CStr(intB))
    Loop Until IsNumeric(Inputopnemen)


End Function

Sub Oef10_6()
    Dim strOut As String
    strOut = CStr(intLeesInteger("Geef een getal op"))
    MsgBox (strOut)
End Sub

Function intLeesInteger(strA) As Integer
    Dim dlNumA As Double

    Do
        dlNumA = CDbl(InputPLS(strA))

    Loop Until -32768 < dlNumA And dlNumA < 32767

    intLeesInteger = CInt(dlNumA)
End Function

Function InputPLS(strA) As String
    Do
        InputPLS = InputBox(strA)
    Loop Until IsNumeric(InputPLS)


End Function

Sub Oef10_7()
    Dim strOut As String
    strOut = CStr(blIntegergetal("Geef een string op"))
    MsgBox (strOut)
End Sub

Function blIntegergetal(strA) As Boolean
    
    Dim strAbc As String
    Dim strOut As String
    Dim intNum As Integer


    strAbc = InputBox(strA)
    If IsNumeric(strAbc) Then
        blIntegergetal = True
    Else
        blIntegergetal = False
    End If
    
End Function

Sub Oef10_8()
    Dim strOut As String
    Dim intNumA As Integer
    For intNumA = 1 To 5
        strOut = CStr(blIntegergetala("Geef een string op", 100 * intNumA))
        MsgBox (strOut)
    Next
End Sub

Function blIntegergetala(strA, intMax) As Boolean
    
    Dim strAbc As String
    Dim strOut As String
    Dim intNum As Integer


    strAbc = InputBox(strA)
    If IsNumeric(strAbc) Then
        If 0 <= CInt(strAbc) And CInt(strAbc) <= intMax Then
            blIntegergetala = True
        Else
            blIntegergetala = False
        End If

    Else
        blIntegergetala = False
    End If
    
End Function

Sub Oef10_9() 'perongeluk al oef 10 ingeboud

    Dim strOut As String
    Dim intNumA As Integer

    For intNumA = 1 To 5
        strOut = CStr(blIntegergetalb("Geef een string op", -100 * intNumA, 100 * intNumA))
        MsgBox (strOut)
    Next
End Sub

Function blIntegergetalb(strA, intMin, intMax) As Boolean
    
    Dim strAbc As String
    Dim strOut As String
    Dim intNum As Integer


    strAbc = InputBox(strA & " tussen " & CStr(intMin) & " en " & CStr(intMax))
    If IsNumeric(strAbc) Then
        If intMin <= CInt(strAbc) And CInt(strAbc) <= intMax Then
            blIntegergetalb = True
        Else
            blIntegergetalb = False
        End If

    Else
        blIntegergetalb = False
    End If
    
End Function

Sub Oef10_10()

    Dim strOut As String
    Dim intNumA As Integer
    Dim strResults As String

    For intNumA = 1 To 5
        strOut = CStr(blIntegergetalb("Geef een string op", -100 * intNumA, 100 * intNumA))
        MsgBox (strOut)
        strResults = strResults & " " & strOut
    Next
    Msgbox(strResults)
    
End Sub





