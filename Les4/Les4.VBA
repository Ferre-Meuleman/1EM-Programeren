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
    Dim bolBoolA As Boolean
    
    strStringA = InputBox("input ")
    


    If VarType(strStringA) = vbInteger  Then
        bolBoolA = True
    Else:
        bolBoolA = False
    End If    
    MsgBox (CStr(bolBoolA))
End Sub
