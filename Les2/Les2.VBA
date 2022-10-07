Option Explicit
'This code is the solution to lesson 2
'Ferre Meuleman
'29/10/2022 

Sub Oef2_1()

    Dim strFirstname As String 'define fristname as a string
    Dim strLastname As String   'define lastname as a string
    Dim strFullname As String   'define fullname as a stirng
    
    strFirstname = InputBox("Vul je voornaamnaam in")
    strLastname = InputBox("Vul je achternaamnaam in")
    
    strFullname = strFirstname & strLastname 'combind the first and last name
    
    
    MsgBox (strFullname) 'Give a popup with the fullname
    
End Sub

Sub Oef2_2()

    Dim strFirstname As String 'define fristname as a string
    Dim strLastname As String   'define lastname as a string
    Dim strFullname As String   'define fullname as a stirng
    
    strFirstname = InputBox("Vul je voornaamnaam in")
    strLastname = InputBox("Vul je achternaamnaam in")
    
    strFullname = "Your name is " & strFirstname & " " & strLastname 'combind the first and last name with a space in between and add your name is to the string
    
    
    MsgBox (strFullname) 'Give a popup with the fullname
    
End Sub

Sub Oef2_3()

    Dim strFirstname As String 'define fristname as a string
    Dim strLastname As String   'define lastname as a string
    
    strFirstname = InputBox("Vul je voornaamnaam in")
    strLastname = InputBox("Vul je achternaamnaam in")
    
    MsgBox ("Your name is " & strFirstname & " " & strLastname) 'Give a popup with firstname and last name combined
    
End Sub


Sub Oef2_4()
    
    MsgBox ("Your name is " & InputBox("Vul je voornaamnaam in") & " " & InputBox("Vul je achternaam in")) 'Give a popup with firstname and last name combined after asking form them in sperated inputs
    
End Sub

Sub Oef2_5()

    Dim intNumA As Integer
    Dim intNumB As Integer
    Dim intSumAB As Integer
    
    intNumA = CInt(InputBox("Number A"))
    intNumB = CInt(InputBox("Number B"))

    intSumAB = intNumA + intNumB

    MsgBox (CStr(intSumAB))
    
End Sub

Sub Oef2_6()

    Dim intNumA As Integer
    Dim intNumB As Integer
    Dim lngSumAB As Long
    
    intNumA = CInt(InputBox("Number A"))
    intNumB = CInt(InputBox("Number B"))

    lngSumAB = CLng(intNumA + intNumB)

    MsgBox (CStr(lngSumAB))
    
End Sub

Sub Oef2_7()

    Dim intNumA As Integer
    Dim intNumB As Integer
    Dim sngSumAB As Single
    
    intNumA = CInt(InputBox("Number A"))
    intNumB = CInt(InputBox("Number B"))

    sngSumAB = CSng(intNumA + intNumB)

    MsgBox (CStr(sngSumAB))
    
End Sub

Sub Oef2_8()

    Dim sngNumA As Single
    Dim sngNumB As Single
    Dim sngSumAB As Single
    
    sngNumA = CSng(InputBox("Number A"))
    sngNumB = CSng(InputBox("Number B"))

    SumAB = CSng(intNumA + intNumB)

    MsgBox (CStr(sngSumAB))
    
End Sub

Sub Oef2_9()

    Dim intNumA As Integer
    Dim intNumB As Integer
    Dim intSumAB As Integer
    
    intNumA = CInt(InputBox("Number A"))
    intNumB = CInt(InputBox("Number B"))

    intSumAB = intNumA + intNumB

    MsgBox (Str(intNumA) & " + " & Str(intNumB) & " = " & CStr(intSumAB))
    
End Sub

Sub Oef2_10()

    Dim intNumA As Integer
    Dim intNumB As Integer
    Dim intDivAB As Integer
    
    intNumA = CInt(InputBox("Number A"))
    intNumB = CInt(InputBox("Number B"))

    intDivAB = intNumA \ intNumB

    MsgBox (Str(intNumA) & " \ " & Str(intNumB) & " = " & CStr(intDivAB))
    
End Sub

Sub Oef2_11()

    Dim intNumA As Integer
    Dim intNumB As Integer
    Dim intDivAB As Integer
    
    intNumA = CInt(InputBox("Number A"))
    intNumB = CInt(InputBox("Number B"))

    intDivAB = intNumA \ intNumB

    MsgBox (Str(intNumA) & " \ " & Str(intNumB) & " = " & CStr(intDivAB))
    MsgBox ("De rest is " & Str(intNumA Mod intNumB))

End Sub

Sub Oef2_12()

    Dim intNumA As Integer
    Dim intNumB As Integer
    Dim intDivAB As Integer
    
    intNumA = CInt(InputBox("Number A"))
    intNumB = CInt(InputBox("Number B"))

    intDivAB = intNumA \ intNumB

    MsgBox (Str(intNumA) & " \ " & Str(intNumB) & " = " & CStr(intDivAB) & vbCrLf & " De rest is " & Str(intNumA Mod intNumB))


End Sub

Sub Oef2_13()

    Dim intNumA As Integer
    Dim intNumB As Integer
    Dim sngDivAB As Single
    
    intNumA = CInt(InputBox("Number A"))
    intNumB = CInt(InputBox("Number B"))

    sngDivAB = CSng(intNumA / intNumB)

    MsgBox (Str(intNumA) & " / " & Str(intNumB) & " = " & CStr(sngDivAB))


End Sub
