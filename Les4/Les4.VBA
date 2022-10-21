Option Explicit
'This code is the solution to lesson 4
'Ferre Meuleman
'13/10/2022

Sub Oef4_1()
    Dim sngNumA As Single
    Dim sngNumB As Single

    sngNumA = CSng(InputBox("input radialen"))
    sngNumB = Cos(sngNumA)

    MsgBox (CStr(sngNumB))
End Sub

Sub Oef4_2()
    Dim sngNumA As Single
    Dim sngNumB As Single

    sngNumA = CSng(InputBox("input graden"))
    sngNumB = Cos(WorksheetFunction.Pi / 180 * sngNumA)

    MsgBox (CStr(sngNumB))
End Sub

Sub Oef4_3()
    Dim sngNumA As Single
    Dim lngNumB As Long

    sngNumA = CSng(InputBox("input "))
    lngNumB = CLng(Fix(Abs(sngNumA)))

    MsgBox (CStr(lngNumB))
End Sub

Sub Oef4_4()
    Dim sngNumA As Single
    Dim lngNumB As Long

    sngNumA = CSng(InputBox("input "))
    lngNumB = Fix(sngNumA)

    MsgBox (CStr(lngNumB))
End Sub

Sub Oef4_5()
    Dim sngNumA As Single
    Dim intNumB As Integer
    Dim lngNumC As Long
    
    sngNumA = CSng(InputBox("input "))
    intNumB = CInt(InputBox("input "))
    lngNumC = Round(sngNumA, intNumB)

    MsgBox (CStr(lngNumC))
End Sub

Sub Oef4_6()
    Dim strStringA As String
    Dim bolBoolA As Boolean
    
    strStringA = InputBox("input ")
    bolBoolA = IsNumeric(strStringA)

    MsgBox (CStr(bolBoolA))

End Sub

Sub Oef4_7()
    Dim strStringA As String    
    
    strStringA = InputBox("input ")

    If ((IsNumeric(strInput) = true) And (Fix(strInput) = strInput) And (-32768 <= strInput <= 32767)) Then
        MsgBox (strStringA)
    Else:
        MsgBox ("foutmelding ")
    End If    
    
End Sub

Sub Oef4_8()
    Dim strStringA As String
    Dim bolBoolA As Boolean
    
    strStringA = InputBox("input ")


    If ((IsNumeric(strInput) = true) And (Fix(strInput) = strInput) And (-32768 <= strInput <= 32767)) Then
        bolBoolA = True
    Else:
        bolBoolA = False
    End If    
    MsgBox (CStr(bolBoolA))
End Sub

Sub Oef4_9()
    Dim sngNumA As Single
    Dim sngNumB As Single
    Dim sngNumC As Single

    sngNumA = Csng(InputBox("input "))
    sngNumB = Csng(InputBox("input "))
    sngNumC = Sqr((sngNumA * sngNumA) + (sngNumB * sngNumB))
    MsgBox (Cstr(sngNumC))
End Sub

Sub Oef4_10()
    Dim sngNumA As Single
    Dim sngNumB As Single
    Dim sngNumC As Single

    sngNumA = Csng(InputBox("R1 "))
    sngNumB = Csng(InputBox("R2 "))
    sngNumC = Csng(InputBox("S "))

    If IsEmpty(sngNumA) Then
        sngNumA = Sqr((sngNumC * sngNumC) - (sngNumB * sngNumB))
    
    ElseIf IsEmpty(sngNumB) Then
        sngNumB = Sqr((sngNumC * sngNumC) - (sngNumA * sngNumA))
    
    ElseIf IsEmpty(sngNumC) Then
        sngNumC = Sqr((sngNumA * sngNumA) + (sngNumB * sngNumB))

    End If

    MsgBox ("R1: "& Cstr(sngNumA) &" R2: "& Cstr(sngNumB) &" R3: "& Cstr(sngNumC))
End Sub

Sub Oef4_11()
    MsgBox "kies", vbYesNo, "kies"
    If Response = vbYes Then
       MsgBox("yes")
    ElseIf Response = vbNo Then 
        MsgBox("No")
    End If
End Sub
