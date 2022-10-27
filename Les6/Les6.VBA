Option Explicit
'This code is the solution to lesson 6
'Ferre Meuleman
'27/10/2022

Sub Oef6_1()
    Dim intArray(4) As Integer 
    Dim intNum As Integer
    For intNum = 0 To 4
        intArray(intNum) = Cint(Inputbox("Input"))
    Next
    For intNum = 0 To 4
        MSgbox("Number" & Cstr(intNum +1 ) & " = " & Cstr(intArray(intNum)))
    Next


End Sub

Sub Oef6_2()
    Dim intArray(4) As Integer 
    Dim intNum As Integer
    Dim strStringA As String
    strStringA = ""
    
    For intNum = 0 To UBound(intArray)
        intArray(intNum) = Cint(Inputbox("Geef een getal voor getal" & Cstr(intNum +1)))
    Next
    
    For intNum = 0 To UBound(intArray)-1
        strStringA = Cstr(strStringA & Cstr(intArray(intNum)) & ", ")
        
    Next
    strStringA = strStringA & Cstr(intArray(4))
    MSgbox(strStringA)

End Sub

Sub Oef6_3()
    Dim intArray(4) As Integer 
    Dim intNum As Integer
    Dim intSum As Integer
    Dim sngAverage As Single
    Dim strStringA As String
    strStringA = ""
    
    For intNum = 0 To UBound(intArray)
        intArray(intNum) = Cint(Inputbox("Geef een getal voor getal" & Cstr(intNum +1)))
        intSum = intSum + intArray(intNum)
    Next
    
    sngAverage = intSum / 5

    For intNum = 0 To UBound(intArray) -1
        strStringA = Cstr(strStringA & Cstr(intArray(intNum)) & ", ") 
    Next

    strStringA = strStringA & Cstr(intArray(4))


    MSgbox "De ingegeven cijfers zijn" & vbCrLf & _
           strStringA & vbCrLf & _
           "Gemiddelde van deze cijfers is" & vbCrLf & _
           Cstr(sngAverage)

End Sub

Sub Oef6_4()
    Dim intArray(14) As Integer 
    Dim strStringA As String
    Dim strStringB As String 
 
    Dim intNum As Integer
    
    Dim intSum As Integer
    Dim intDiv As Integer
    Dim sngAverage As Single
    
    strStringA = ""
    
    For intNum = 0 To UBound(intArray)
        strStringA = Inputbox("Geef een getal voor getal" & Cstr(intNum +1))
        if StrPtr(strStringA) = 0 Then
            intArray(intNum) = Cint(strStringA)
            intSum = intSum + intArray(intNum)
            intDiv = intDiv + 1 
        End If
    Next
    MSgbox "Gemiddelde word berekent"
    sngAverage = intSum / intDiv

    For intNum = 0 To UBound(intArray) -1
        strStringA = Cstr(intArray(intNum))
        if StrPtr(strStringA) = 0 Then
            strStringB = Cstr(strStringB & Cstr(intArray(intNum)) & ", ") 
        End If
    Next
    strStringB = Cstr(strStringA & Cstr(intArray(intNum))) 

    MSgbox "De ingegeven cijfers zijn" & vbCrLf & _
            strStringB & vbCrLf & _
            "Gemiddelde van deze cijfers is" & vbCrLf & _
            Cstr(sngAverage)

End Sub

Sub Oef6_5()
    Dim intArray(4) As Integer 
    Dim intNum As Integer
    Dim intSmall As Integer
    
    For intNum = 0 To UBound(intArray)
        intArray(intNum) = Inputbox("Geef een getal voor getal" & Cstr(intNum +1))
    Next

    intSmall = intArray(0)

    For intNum = 1 To UBound(intArray) 
        if intArray(intNum) < intSmall then
            intSmall = intArray(intNum)
        End if
    Next

    MSgbox "Kleinste getal is " & Cstr(intSmall)
    
End Sub

Sub Oef6_6()
    Dim intArray(9) As Integer 
    Dim intNum As Integer
    Dim intSmall As Integer
    
    For intNum = 0 To UBound(intArray)
        intArray(intNum) = Inputbox("Geef een getal voor getal" & Cstr(intNum +1))
    Next

    intSmall = intArray(0)

    For intNum = 1 To UBound(intArray) 
        if intArray(intNum) < intSmall then
            intSmall = intArray(intNum)
        End if
    Next

    MSgbox "Kleinste getal is " & Cstr(intSmall)
    
End Sub

Sub Oef6_7()
    Dim intArray(4) As Integer
    Dim intNum As Integer
    Dim intBig As Integer
    Dim intBigAmount As Integer
    
    For intNum = 0 To UBound(intArray)
        intArray(intNum) = Inputbox("Geef een getal voor getal" & Cstr(intNum +1))
    Next

    intBig = intArray(0)

    For intNum = 1 To UBound(intArray) 
        if intArray(intNum) >= intBig then
            intBig = intArray(intNum)
            intBigAmount = intBigAmount +1
        End if
    Next
    if intArray(0) = intBig then
        intBigAmount = intBigAmount +1
    End if
    MSgbox "Kleinste getal is " & Cstr(intBig) & " en komt "& Cstr(intBigAmount) & " keer voor"
    
End Sub

Sub Oef6_8()
    Dim intArray(4) As Integer
    Dim intNum As Integer
    Dim intBig As Integer
    Dim intBigIndex As Integer
    For intNum = 0 To UBound(intArray)
        intArray(intNum) = Inputbox("Geef een getal voor getal" & Cstr(intNum +1))
    Next

    intBig = intArray(0)

    For intNum = 1 To UBound(intArray) 
        if intArray(intNum) > intBig then
            intBig = intArray(intNum)
            intBigIndex = intNum
        End if
    Next

    MSgbox "Kleinste getal is " & Cstr(intBig) & " op locatie "& Cstr(intBigIndex)
    
End Sub

Sub Oef6_9()
    Dim intArray(3, 3) As Integer

    Dim intRow As Integer
    Dim intCol As Integer
    
    Dim strStringA As String

    For intRow = 0 To 3
        For intCol = 0 To 3
        intArray(intRow, intCol) = InputBox("Geef een getal voor getal" & CStr(intRow + 1) & "," & CStr(intCol + 1))
        Next
    Next

    For intRow = 0 To 3
        For intCol = 0 To 3
            strStringA = strStringA & intArray(intRow, intCol) & vbTab
        Next 
        strStringA = strStringA & vbCrLf
    Next 

    MsgBox(strStringA)

End Sub

Sub Oef6_10()
    Dim intArray(3,3) As Integer
    Dim strStringA As String
    Dim intRow As Integer
    Dim intCol As Integer

    
    For intRow = 0 to 3
        For intCol = 0 To 3
            
            strStringA = Inputbox("Geef een getal voor getal" & Cstr(intRow +1) &"," &Cstr(IntCol +1))
            if IsNumeric(strStringA) = False Then
                do 
                    strStringA = Inputbox("Geef een getal voor getal" & Cstr(intRow +1) &"," &Cstr(IntCol +1))
                Loop Until IsNumeric(strStringA) = True 
            Elseif Cint(strStringA) < 0 or Cint(strStringA) > 9 then
                do 
                    strStringA = Inputbox("Geef een getal voor getal" & Cstr(intRow +1) &"," &Cstr(IntCol +1))
                Loop Until Cint(strStringA) >= 0 And Cint(strStringA) <= 9 
            End IF
            intArray(intRow, intCol) = Cint(strStringA)        
        Next
    Next

    For intRow = 0 To 3
        For intCol = 0 To 3
            strStringA = strStringA & intArray(intRow, intCol) & vbTab
        Next 
        strStringA = strStringA & vbCrLf
    Next 

    MsgBox(strStringA)

    
End Sub

