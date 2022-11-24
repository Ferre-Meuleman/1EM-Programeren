Option Explicit
'This code is the solution to lesson 9
'Ferre Meuleman
'24/11/2022

Dim intNumA As Integer 'Globle variable nodig voor oef 10


Sub Oef9_1()
    Dim intNumA As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strSheet As String
    
    intNumA = Cint(InputBox("Getal"))
    intRow = CInt(InputBox("Rij"))
    intCol = CInt(InputBox("Col"))
    strSheet = InputBox("Sheet")

    Worksheets(strSheet).Activate
    Cells(intRow, intCol).Value = intNumA

End Sub

Sub Oef9_2()
    Dim intNumA As Integer
    Dim intRow As Integer
    Dim strCol As String
    Dim strSheet As String
    
    intNumA = InputBox("Getal")
    intRow = CInt(InputBox("Rij"))
    strCol = InputBox("Col")
    strSheet = InputBox("Sheet")

    Worksheets(strSheet).Activate
    Cells(intRow, strCol).Value = intNumA

End Sub

Sub Oef9_3()
    Dim strArray(4) As String
    Dim i As Integer
    
    For i = 0 to UBound(strArray) 
        strArray(i) = InputBox("Value for Cel: " & Cstr(i+1))
    Next 

    For i= 0 to UBound(strArray)
        Cells(6 + i ,"B").Value = strArray(i)
    Next

End Sub

Sub Oef9_4()
    Dim strArray(4) As String
    Dim i As Integer
    For i = 0 to UBound(strArray) 
        strArray(i) = Cells(5 + i ,"C").Value
        Cells(5 + i ,"C").Clear

    Next 

    For i= 0 to UBound(strArray)
        Cells(10,"B").Offset(,i).Value = strArray(i)
    Next

End Sub

Sub Oef9_5() 'plak deze code in de sub van de aangamakt knop
    Dim strArray(4) As String
    Dim i As Integer
    For i = 0 to UBound(strArray) 
        strArray(i) = Cells(5 + i ,"C").Value
        Cells(5 + i ,"C").Clear
    Next 

    For i= 0 to UBound(strArray)
        Cells(10,"B").Offset(,i).Value = strArray(i)
    Next

End Sub

Sub Oef9_6() 'plak deze code in de sub van de aangamakt knop

    ActiveCell.Offset(,1).Value = Cint(ActiveCell.Value)*2
    
End Sub

Sub oef9_7()
    Dim sngGetal As Single
    sngGetal = CSng(InputBox("Geef getal:"))
    If sngGetal >= 0 Then
        Call PrintPositief
    Else
        Call PrintNegatief
    End If
    sngGetal = CSng(InputBox("Geef nog een getal:"))
    If sngGetal >= 0 Then
        Call PrintPositief
    Else
       Call PrintNegatief
    End If
    sngGetal = CSng(InputBox("Geef laatste getal:"))
    If sngGetal >= 0 Then
        Call PrintPositief
    Else
       Call PrintNegatief
    End If
End Sub

Sub PrintPositief()
    MsgBox ("Getal is positief")
    MsgBox ("U doet dus winst")

End Sub

Sub PrintNegatief()
    MsgBox ("Getal is negatief")
    MsgBox ("U doet dus verlies")

End Sub

Sub oef9_8()
    Dim sngGetal As Single
    sngGetal = CSng(InputBox("Geef getal:"))
    If sngGetal >= 0 Then
      Call  PrintBoodschapPositief (sngGetal)
    Else
        Call PrintBoodschapNegatief (sngGetal)
    End If
    sngGetal = CSng(InputBox("Geef nog een getal:"))
    If sngGetal >= 0 Then
       Call PrintBoodschapPositief (sngGetal)
    Else
       Call PrintBoodschapNegatief (sngGetal)
    End If
    sngGetal = CSng(InputBox("Geef laatste getal:"))
    If sngGetal >= 0 Then
       Call PrintBoodschapPositief (sngGetal)
    Else
       Call PrintBoodschapNegatief (sngGetal)
    End If
End Sub

Sub PrintBoodschapPositief(sngGetal)
    MsgBox ("Getal " & CStr(sngGetal) & " is positief")
    MsgBox ("U doet dus winst")
End Sub

Sub PrintBoodschapNegatief(sngGetal)
    MsgBox ("Getal " & CStr(sngGetal) & " is negatief")
    MsgBox ("U doet dus verlies")

End Sub



Sub Oef9_9()
    Dim strVal As String
    Dim intRow As Integer
    Dim intCol As String

    strVal = InputBox("Value")
    intRow = CInt(InputBox("Rij"))
    intCol = CInt(InputBox("Col"))

    Call Oef9_9_1(strVal, intRow, intCol)

End Sub


Sub Oef9_9_1(strVal, intRow, intCol)
    
    Cells(intRow, intCol).Value = strVal

End Sub

Sub Oef9_10()
    

    Call Oef9_10_1()
    Call Oef9_10_2()

End Sub


Sub Oef9_10_1()
    
    IntNumA = CInt(InputBox("Getal"))
    
End Sub

Sub Oef9_10_2()
    
   MsgBox("Het getal is "& Cstr(IntNumA))
    
End Sub