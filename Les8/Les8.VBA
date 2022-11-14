Option Explicit
'This code is the solution to lesson 8
'Ferre Meuleman
'27/10/2022


Sub Oef8_1()
    Dim intArray(3, 3) As Integer
    Dim strStringA As String
    Dim intRow As Integer
    Dim intCol As Integer
    Dim intSmall As Integer

    
    For intRow = 0 To 3
        For intCol = 0 To 3
            
            strStringA = InputBox("Geef een getal voor getal" & CStr(intRow + 1) & "," & CStr(intCol + 1))
            If IsNumeric(strStringA) = False Then
                Do
                    strStringA = InputBox("Geef een getal voor getal" & CStr(intRow + 1) & "," & CStr(intCol + 1))
                Loop Until IsNumeric(strStringA) = True
            ElseIf CInt(strStringA) < 0 Or CInt(strStringA) > 100 Then
                Do
                    strStringA = InputBox("Geef een getal voor getal" & CStr(intRow + 1) & "," & CStr(intCol + 1))
                Loop Until CInt(strStringA) >= 0 And CInt(strStringA) <= 100
            End If
            intArray(intRow, intCol) = CInt(strStringA)
        Next
    Next

    intSmall = intArray(0, 0)
    For intRow = 0 To 3
        For intCol = 0 To 3
            If intSmall > intArray(intRow, intCol) Then
                intSmall = intArray(intRow, intCol)
            End If
        Next
    Next

    MsgBox (CStr(intSmall))

    
End Sub

Sub Oef8_2()
    Dim intArray(3, 1) As Integer
    Dim strStringA As String
    Dim intRow As Integer
    Dim intCol As Integer
    Dim intSmall As Integer

    
    For intRow = 0 To 3
        For intCol = 0 To 1
            
            strStringA = InputBox("Geef een getal voor getal" & CStr(intRow + 1) & "," & CStr(intCol + 1))
            If IsNumeric(strStringA) = False Then
                Do
                    strStringA = InputBox("Geef een getal voor getal" & CStr(intRow + 1) & "," & CStr(intCol + 1))
                Loop Until IsNumeric(strStringA) = True
            ElseIf CInt(strStringA) < 0 Or CInt(strStringA) > 100 Then
                Do
                    strStringA = InputBox("Geef een getal voor getal" & CStr(intRow + 1) & "," & CStr(intCol + 1))
                Loop Until CInt(strStringA) >= 0 And CInt(strStringA) <= 100
            End If
            intArray(intRow, intCol) = CInt(strStringA)
        Next
    Next

    intSmall = intArray(0, 0)
    For intRow = 0 To 3
        For intCol = 0 To 1
            If intSmall > intArray(intRow, intCol) Then
                intSmall = intArray(intRow, intCol)
            End If
        Next
    Next

    MsgBox (CStr(intSmall))

    
End Sub

Sub Oef8_3()
    Dim intArray(3, 3) As Integer
    Dim strStringA As String
    Dim intRow As Integer
    Dim intCol As Integer
    Dim intBig As Integer
    Dim intBigCount As Integer
    
    For intRow = 0 To 3
        For intCol = 0 To 3
            
            strStringA = InputBox("Geef een getal voor getal" & CStr(intRow + 1) & "," & CStr(intCol + 1))
            If IsNumeric(strStringA) = False Then
                Do
                    strStringA = InputBox("Geef een getal voor getal" & CStr(intRow + 1) & "," & CStr(intCol + 1))
                Loop Until IsNumeric(strStringA) = True
            ElseIf CInt(strStringA) < 0 Or CInt(strStringA) > 100 Then
                Do
                    strStringA = InputBox("Geef een getal voor getal" & CStr(intRow + 1) & "," & CStr(intCol + 1))
                Loop Until CInt(strStringA) >= 0 And CInt(strStringA) <= 100
            End If
            intArray(intRow, intCol) = CInt(strStringA)
        Next
    Next

    intBig = intArray(0, 0)
    For intRow = 0 To 3
        For intCol = 0 To 3
            If intBig >= intArray(intRow, intCol) Then
                intBig = intArray(intRow, intCol)
                intBigCount = intBigCount + 1
            End If
        Next
    Next
    intBigCount = intBigCount - 1
    MsgBox ("Het grootste cijfer is: " & CStr(intBig) & "en het komt " & CStr(intBigCount) & "keer voor")

    
End Sub

Sub Oef8_4()
    Dim intArrayA(4) As Integer
    Dim intArrayB(4) As Integer
    Dim strStringA As String
    Dim intNumA As Integer
    Dim intNumB As Integer
    Dim intNumC As Integer
    Dim intResult As Integer

    For intNumA = 0 To 4
        
        strStringA = InputBox("Geef een getal voor getal A:" & CStr(intNumA + 1))
        If IsNumeric(strStringA) = False Then
            Do
                strStringA = InputBox("Geef een getal voor getal A:" & CStr(intNumA + 1))
            Loop Until IsNumeric(strStringA) = True
        ElseIf CInt(strStringA) < 0 Or CInt(strStringA) > 100 Then
            Do
                strStringA = InputBox("Geef een getal voor getal A:" & CStr(intNumA + 1))
            Loop Until CInt(strStringA) >= 0 And CInt(strStringA) <= 100
        End If
        intArrayA(intNumA) = CInt(strStringA)
    Next

    For intNumB = 0 To 4
        
        strStringA = InputBox("Geef een getal voor getal B:" & CStr(intNumB + 1))
        If IsNumeric(strStringA) = False Then
            Do
                strStringA = InputBox("Geef een getal voor getal B:" & CStr(intNumB + 1))
            Loop Until IsNumeric(strStringA) = True
        ElseIf CInt(strStringA) < 0 Or CInt(strStringA) > 100 Then
            Do
                strStringA = InputBox("Geef een getal voor getal B:" & CStr(intNumB + 1))
            Loop Until CInt(strStringA) >= 0 And CInt(strStringA) <= 100
        End If
        intArrayB(intNumB) = CInt(strStringA)
    Next

    strStringA = ""
    For intNumC = 0 To 3
        strStringA = strStringA & CStr(intArrayA(intNumC) * intArrayB(intNumC)) & "+"
        intResult = intResult + (intArrayA(intNumC) * intArrayB(intNumC))
    Next
    intResult = intResult + (intArrayA(4) * intArrayB(4))
    strStringA = strStringA & CStr(intArrayA(4) * intArrayB(4)) & "=" & CStr(intResult)
    MsgBox (strStringA)

    
End Sub

Sub Oef8_5()

    Dim intArrayA(14, 1) As Integer
    Dim intNumA As Integer
    Dim strStringA As String
    Dim Sum As Integer
    Dim Average As Integer
    Dim intNumB As Integer
    Dim intTemp As Integer

    Do
        For intNumA = LBound(intArrayA) To UBound(intArrayA)
            If intArrayA(intNumA, 1) = 0 Then
                strStringA = InputBox("Geef een getal voor getal A:" & CStr(intNumA + 1))
            
                If strStringA <> vbNullString Then
                    intArrayA(intNumA, 1) = 1
                    intArrayA(intNumA, 0) = CInt(strStringA)
                End If
                
            End If
        Next

        Sum = 0

        For intNumA = LBound(intArrayA) To UBound(intArrayA)
            Sum = Sum + intArrayA(intNumA, 1)
        Next

    Loop Until Sum = 15

    MsgBox "Je hebt 15 getallen ingegeven het gemiddelde word berekent"

    For intNumA = LBound(intArrayA) To UBound(intArrayA)
        Average = Average + intArrayA(intNumA,0)
    Next

    Average = Average / 15
    
    intNumB = 13
    do     
        For intNumA = 0 to Ubound(intArrayA)-1 
            if intArrayA(intNumA,0) > intArrayA(intNumA +1,0) Then
                intTemp = intArrayA(intNumA,0)
                intArrayA(intNumA,0) = intArrayA(intNumA +1,0)
                intArrayA(intNumA +1,0) = intTemp
            End if
        Next
        intNumB = intNumB - 1
    
    loop until intNumB = 0

    strStringA = ""
    For intNumA = 0 To UBound(intArrayA)-1
        strStringA = Cstr(strStringA & Cstr(intArray(intNumA)) & ", ")
        
    Next
    strStringA = strStringA + Cstr(intArrayA(14)) 

    MsgBox "Het gemiddelde is: " & Cstr(Average) & vbCrLf & "De array gesorteert van klein naar groot is:" & vbCrLf & strStringA
End Sub

'Voor oef 6-12 zie onder Fromulieren 

    
