Option Explicit
'This code is the solution to lesson 3
'Ferre Meuleman
'6/11/2022

Sub Oef3_1()

    Dim intNumA As Integer
    Dim strAwns As String

    intNumA = CInt(InputBox("Number A"))
    If intNumA Mod 2 = 0 Then
        strAwns = Str(intNumA) + " is divisible by 2"
    End If
    MsgBox (strAwns)

End Sub

Sub Oef3_2()

    Dim intNumA As Integer
    Dim strAwns As String

    intNumA = CInt(InputBox("Number A"))
    If intNumA Mod 2 = 0 Then
        strAwns = Str(intNumA) + " is divisible by 2"

    Else:
        strAwns = Str(intNumA) + " is not divisible by 2"
    
    End If
    MsgBox (strAwns)

End Sub

Sub Oef3_3()

    Dim intNumA As Single
    Dim strAwns As String

    intNumA = CInt(InputBox("Number A"))
    
    If intNumA > 0 Then
        strAwns = Str(intNumA) + " is bigger than 0"
    
    ElseIf intNumA < 0 Then
        strAwns = Str(intNumA) + " is smaller than 0"
    
    Else:
        strAwns = Str(intNumA) + " is 0"
    
    End If
    MsgBox (strAwns)

End Sub

Sub Oef3_4()

    Dim intNumA As Single
    Dim strAwns As String

    intNumA = CInt(InputBox("Number A"))
    If intNumA Mod 2 = 0 Then
        strAwns = Str(intNumA) + " is divisible by 2"

    Else:
        strAwns = Str(intNumA) + " is not divisible by 2"
    
    End If
    MsgBox (strAwns)

End Sub

Sub Oef3_5()
    Dim strLetter As String
    Dim strAwns As String

    strLetter = LCase(InputBox("Input a letter"))

    If strLetter < "k" Then
        strAwns = strLetter + " comes before k in the alfabet"
    
    ElseIf strLetter > "k" Then
        strAwns = strLetter + " comes after k in the alfabet"
    
    Else:
        strAwns = strLetter + " is the letter k"

    End If
    MsgBox (strAwns)

End Sub

Sub Oef3_6()

    Dim boolA As Boolean

    boolA = True
    MsgBox (boolA)
    boolA = False
    MsgBox (boolA)

End Sub

Sub Oef3_7()

    Dim boolA As Boolean

    boolA = InputBox("input boolean")
    MsgBox (boolA)

End Sub

Sub Oef3_8()

    Dim boolA As Boolean
    Dim boolB As Boolean
    Dim boolC As Boolean

    boolA = InputBox("input boolean A")
    boolB = InputBox("input boolean B")
    
    boolC = Not boolA And Not boolB


    
    MsgBox (boolC)

End Sub

Sub Oef3_9()

    Dim boolA As Boolean
    Dim boolB As Boolean
    Dim boolC As Boolean

    boolA = InputBox("input boolean A")
    boolB = InputBox("input boolean B")
    
    boolC = boolA And boolB
    
    MsgBox (boolC)

End Sub

Sub Oef3_10()

    Dim boolA As Boolean
    Dim boolB As Boolean
    Dim boolC As Boolean

    boolA = InputBox("input boolean A")
    boolB = InputBox("input boolean B")
    
    boolC = boolA = Not boolB
    
    MsgBox (boolC)

End Sub

Sub Oef3_11()

    Dim boolA As Boolean
    Dim boolB As Boolean
    Dim boolC As Boolean

    boolA = InputBox("input boolean A")
    boolB = InputBox("input boolean B")
    
    boolC = Not BoolA and not boolB
    
    MsgBox (boolC)

End Sub

Sub Oef3_12()

    Dim boolA As Boolean

    boolA = CBool(InputBox("input boolean A"))
    if boolA = True Then
        MsgBox ("You enterd" + str(boolA))
    else:
        MsgBox ("You enterd" + str(boolA))
    End if

End Sub

Sub Oef3_13()

    Dim sinNumA As Single
    Dim sinNumB As Single
   
    sinNumA = CSng(InputBox("input number"))
    sinNumB = CSng(InputBox("input number"))

    if sinNumA < sinNumB then
        sinNumB = CSng(InputBox("input number"))
    elseif sinNumA > sinNumB then
        sinNumA = CSng(InputBox("input number"))
    else:
        sinNumA = Csng(InputBox("input number"))
    End if 

    if sinNumA < sinNumB then
        sinNumB = CSng(InputBox("input number"))
    elseif sinNumA > sinNumB then
        sinNumA = CSng(InputBox("input number"))
    else:
        sinNumA = Csng(InputBox("input number"))
    End if

    if sinNumA < sinNumB then
        sinNumB = CSng(InputBox("input number"))
    elseif sinNumA > sinNumB then
        sinNumA = CSng(InputBox("input number"))
    else:
        sinNumA = Csng(InputBox("input number"))
    End if

    if sinNumA < sinNumB then
        sinNumB = CSng(InputBox("input number"))
    elseif sinNumA > sinNumB then
        sinNumA = CSng(InputBox("input number"))
    else:
        sinNumA = Csng(InputBox("input number"))
    End if
    
    if sinNumA < sinNumB then
        Msgbox(Cstr(sinNumA))
    else:
        Msgbox(Cstr(sinNumB))
    End if
End Sub
