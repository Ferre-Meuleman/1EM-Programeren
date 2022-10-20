Option Explicit
'This code is the solution to lesson 4
'Ferre Meuleman
'13/10/2022

Sub Oef5_1()

    Dim intA As Integer

    For intA = 1 To 5
        MsgBox (intA)
    
    Next intA
End Sub

Sub Oef5_2()

    Dim intA As Integer

    For intA = 1 To 5
        MsgBox (intA)
        intA = intA + 1
    Next intA
End Sub

Sub Oef5_3()

    Dim intA As Integer
    intA = 2
    Do
        MsgBox (intA * 3)
        intA = intA + 1
    Loop While intA * 3 < 26
End Sub

Sub Oef5_4()

    Dim intA As Integer
    intA = 2
    Do Until intA * 4 >= 25
        MsgBox (intA * 4)
        intA = intA + 1
    Loop
End Sub

Sub Oef5_5()

    Dim sngNumA As Single
    sngNumA = 2.5
    
    Do While sngNumA <= 20
        MsgBox (sngNumA)
        sngNumA = sngNumA + 2.5
    Loop
End Sub

Sub Oef5_6()

    Dim sngNumA As Single
    sngNumA = 4
    
    Do Until sngNumA >= 7
        MsgBox (sngNumA)
        sngNumA = sngNumA + 0.5
    Loop
End Sub

Sub Oef5_7()

    Dim intMax As Integer
    Dim intSom As Integer
    Dim intNumA As Integer

    sngNumA = 4
    intMax = Cint(InputBox ("maximum"))
    
    Do While intsom <= intmax
        intNumA = intNumA +1
        intSom = intSom + intNumA
    Loop
    intSom = intSom - intNumA
    MsgBox("De som 1+ ..." & intNumA &" is " &intSom)
End Sub


Sub Oef5_8()
    Dim strInput As String
       
    Do
        strInput = InputBox("maximum")
        
    Loop Until IsNumeric(strInput) = True
    MsgBox (strInput)
End Sub

Sub Oef5_9()
    Dim strInput As String

    strInput = InputBox("maximum")

    Do Until IsNumeric(strInput) = True
        strInput = InputBox("Was geen getal")
    Loop

    MsgBox (strInput)
End Sub

Sub Oef5_10()
    Dim strInput As String
    Dim intPoging As Integer

    strInput = InputBox("maximum")

    Do Until (IsNumeric(strInput) = True) Or (intPoging > 4)
        
        strInput = InputBox("Was geen getal, u heeft nog " & Str(5 - intPoging) & "pogingen over")
        intPoging = intPoging + 1
        
        If intPoging = 5 Then
            strInput = "U pogingen zijn op"
        End If

    Loop

    MsgBox (strInput)

End Sub

Sub Oef5_11()
    Dim intPoging As Integer
    Dim intPogingMax As Integer
    Dim sngNumA As Single
    Dim sngNumB As Single
    intPogingMax = 5

    For intPoging = 1 To intPogingMax
        sngNumA = CSng(InputBox("input"))
        If sngNumA > sngNumA Then
            sngNumB = sngNumA
        End If
    Next
    MsgBox (sngNumB)

End Sub

Sub Oef5_12()
    Dim intPoging As Integer
    Dim intPogingMax As Integer
    Dim strInput As String
    Dim sngNumA As Single
    Dim sngNumB As Single

    intPogingMax = 10

    
    For intPoging = 1 To intPogingMax
        
        Do
            strInput = (InputBox("input "))
        Loop Until (IsNumeric(strInput) = True)
        
        sngNumA = strInput
        
        If sngNumA > sngNumB Then
            sngNumB = sngNumA
        End If
    Next
    MsgBox (sngNumB)

End Sub

Sub Oef5_13()
    Dim strInput As String
    strInput = (InputBox("input "))
    do 
    
        If IsNumeric(strInput) = False Then
            strInput = (InputBox("input een getal"))
    
        ElseIf (Fix(strInput) <> strInput) Then
            strInput = (InputBox("input een geheel getal"))
    
        ElseIf (-32.768 > strInput Or strInput > 32.767) Then
            strInput = (InputBox("input een getal tussen de waarden van een int"))
    
        End If
    loop Until ((IsNumeric(strInput) = true) And (Fix(strInput) = strInput) And (-32768 < strInput < 32767))
    
    MsgBox (strInput)
End Sub

