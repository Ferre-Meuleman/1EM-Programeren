Option Explicit
'This code is the solution to lesson 11
'Ferre Meuleman
'8/12/2022

Sub Oef11_1()
    Dim intA As Integer
    Dim intB As Integer

    intA = 10
    intB = 20

    MsgBox ("IntA voor Swap: " & CStr(intA))
    MsgBox ("Intb voor Swap: " & CStr(intB))

    Call WisselInt(intA, intB)

    MsgBox ("IntA na Swap: " & CStr(intA))
    MsgBox ("Intb na Swap: " & CStr(intB))

End Sub


Sub WisselInt(ByRef intA As Integer, ByRef intB As Integer)
    Dim intC As Integer
    
    intC = intA
    intA = intB
    intB = intC

End Sub

Sub Oe11_2()
    Const Aantal = 5
    Dim strGegevens(Aantal) As String
    Dim intI As Integer
    For intI = 1 To Aantal
        strGegevens(intI) = "Testgegevens" & CStr(intI)
    Next intI
    Call Array_to_excel(strGegevens(), Aantal)
End Sub

Sub Array_to_excel(strGegevens() As String, Aantal As Integer)
    Dim intI As Integer
    For intI = 1 To Aantal
        Cells(intI, 2).value = strGegevens(intI)
    Next intI
    
End Sub

Sub Oe11_3()
    Const Aantal = 5
    Const Col = 1
    Dim strGegevens(Aantal) As String
    Dim intI As Integer
    For intI = 1 To Aantal
        strGegevens(intI) = "Testgegevens" & CStr(intI)
    Next intI
    Call Array_to_excelB(strGegevens(), Aantal, "De waarde voor string", Col)
End Sub

Sub Array_to_excelB(strGegevens() As String, Aantal As Integer, Melding As String, Col As Integer)
    Dim intI As Integer
    For intI = 1 To Aantal
        Cells(intI, Col).value = Melding & CStr(intI) & ":"
        Cells(intI, 2).value = strGegevens(intI)
    Next intI
    
End Sub

Sub Oe11_4()
    Const Aantal = 5
    Const Col = 1
    Dim strGegevens(Aantal) As String
    Dim intI As Integer
    Call Gegevensinlees(strGegevens(), Aantal)
    
    Call Array_to_excelB(strGegevens(), Aantal, "De waarde voor string", Col)
End Sub

Sub Gegevensinlees(ByRef strGegevens() As String, Aantal As Integer)
    Dim intI
    For intI = 1 To Aantal
        strGegevens(intI) = InputBox("Geef een waarde in voor " & CStr(intI))
    Next intI

End Sub

Sub Oe11_5()
    Const Aantal = 5
    Const Col = 1
    Dim strGegevens(Aantal) As String
    Dim intA As Integer

    Call GegevensinleesB(strGegevens(), Aantal, intA)
    
    Call Array_to_excelB(strGegevens(), intA, "De waarde voor string", Col)
End Sub

Sub GegevensinleesB(ByRef strGegevens() As String, Aantal As Integer, ByRef intA As Integer)
    Dim intI As Integer
    intI = 1
    Do
        strGegevens(intI) = InputBox("Geef een waarde in voor " & CStr(intI))
        intI = intI + 1
    Loop Until intI = Aantal Or strGegevens(intI - 1) = vbNullString
    intA = intI - 2


End Sub

Sub Oe11_6()
    Const intA = 5
    Dim intInput(intA) As Integer
    Dim intOutput(intA) As Integer
    Dim strInput As String
    Dim strOutput As String
    Dim intI As Integer

    For intI = 0 To intA
        intInput(intI) = CInt(InputBox("Getal " & CStr(intI)))
        strInput = strInput & CStr(intInput(intI)) & " "
    Next intI
    
    Call verdubbel(intA, intInput, intOutput)
    
    

    MsgBox ("Input: " & vbCrLf & strInput & vbCrLf & "Output: " & strOutput)
End Sub

Sub verdubbel(intA As Integer, ByRef intInput() As Integer, ByRef intOutput() As Integer)
    Dim intI As Integer

    For intI = 0 To intA
        intOutput(intI) = intInput(intI) * 2
    Next intI
    
End Sub


Sub Oe11_7()
    Const intA = 4
    Dim strInput(intA) As String
    Dim intOutput(intA) As Integer
    Dim strOutput As String
    Dim intI As Integer

    For intI = 0 To intA
        do 
            strInput(intI) = InputBox("Getal " & CStr(intI-1))
        Loop until  Checknumber(strInput(intI)) = True
    Next intI

    Call verdubbelb(intA, strInput, intOutput)

    For intI = 0 To intA
        strOutput = strOutput & CStr(intOutput(intI)) & " "
    Next intI
    
    MsgBox ("Output: " & vbCrLf & strOutput)
End Sub

Function Checknumber(strA As String) As Boolean
    if IsNumeric(strA) Then
        If CDbl(strA) < 500 and CDbl(strA) > -500 then
            if Fix(Cint(strA)) = Cint(strA) then 
                Checknumber = True
            End if 
        End if 
    Else 
        Checknumber = False
    End if 
End Function

Sub verdubbelb(intA As Integer, ByRef strInput() As String, ByRef intOutput() As Integer)
    Dim intI As Integer

    For intI = 0 To intA
        intOutput(intI) = Cint(intInput(intI)) * 2
    Next intI
    
End Sub
