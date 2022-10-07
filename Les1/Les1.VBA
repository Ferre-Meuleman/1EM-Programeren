Option Explicit
'This code is the solution to lesson 1
'Ferre Meuleman
'22/10/2022

Sub Oef1_1()
    MsgBox ("Ferre")
End Sub

Sub Oef1_2()
    Dim strNaam As String
    strNaam = "Ferre"
    MsgBox (strNaam)
    
End Sub

Sub Oef1_3()
    Const Naam As String = "Ferre"
    MsgBox (Naam)
    
End Sub

Sub Oef1_4()
    Dim strNaam As String
    strNaam = InputBox("Vul je naam in")
    MsgBox (strNaam)
    
End Sub


Sub Oef1_5()

    Const intGeboortejaar As Integer = 2004
    Dim strTussenstap As String
    strTussenstap = CStr(intGeboortejaar)
    MsgBox (strTussenstap)
    
End Sub

Sub Oef1_6()

    Const intGeboortejaar As Integer = 2004
    MsgBox (CStr(intGeboortejaar))
    
End Sub


Sub Oef1_7()

    Dim intGeboortejaar As Integer
    intGeboortejaar = CInt(InputBox("geef een integer getal in"))
    MsgBox (intGeboortejaar)
    
End Sub
Sub Oef1_8()

    Dim lngtGeboortejaar As Long
    lngGeboortejaar = CLng(InputBox("geef een integer getal in"))
    MsgBox (lngGeboortejaar)
    
End Sub
Sub Oef1_9()

    Dim sngGeboortejaar As Single
    sngGeboortejaar = CSng(InputBox("geef een integer getal in"))
    MsgBox (sngGeboortejaar)
    
End Sub
Sub Oef1_10()

    Dim dblGeboortejaar As Double
    dblGeboortejaar = CDbl(InputBox("geef een integer getal in"))
    MsgBox (dblGeboortejaar)
    
End Sub
