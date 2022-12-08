# Basisoefeningen les 10
## Oefening 1
Geef volgend programma in en test het:
```VBA
Sub Oefening1()
    Dim intA As Integer, intB As Integer
    
    intA = CInt(InputBox("Geef een integergetal in"))
    intB= intA \ 2
    
    MsgBox ("Het resultaat is " & CStr(intB))

End Sub
```
Pas het programma aan door het bijschrijven van een functie zodat het hoofdprogramma er als volgt uit ziet:
(De werking van het programma wijzigt niet.)

```VBA
Sub Oefening1()
    Dim intA As Integer, intB As Integer
        
    intA = CInt(InputBox("Geef een integergetal in"))
    intB = intDubbel(intA)
    
    MsgBox ("Het resultaat is " & CStr(intB))
End Sub
```

## Oefening 2
Pas het programma van oefening 1 aan door het bijschrijven van een functie zodat het hoofdprogramma er als volgt uit ziet:
(De werking van het programma wijzigt niet.)

```VBA
Sub Oefening2()
    Dim intA As Integer, intB As Integer

    intA = intlees()
    intB = intDubbel(intA)

    MsgBox ("Het resultaat is " & CStr(intB))
End Sub
```

## Oefening 3
Pas het programma van oefening 2 aan door het bijschrijven van een subroutine zodat het hoofdprogramma er als volgt uit ziet:
(De werking van het programma wijzigt niet.)


```VBA
Sub Oefening3()
    Dim intA As Integer, intB As Integer
    intA = intlees()
    intB = intDubbel(intA)
    
    Call drukAf(intB)
End Sub
```
## Oefening 4
Geef volgend programma in en test het:


```VBA
Sub oef10_4()
    Const MIN = 5
    Const MAX = 15

    Dim intA As Integer

    intA = CInt(InputBox("Geef getal"))

    If intA <= MAX And intA >= MIN Then
        MsgBox ("OK")
    Else
        MsgBox ("NOK")
    End If

End Sub

```



Pas het programma aan door het bijschrijven van een functie zodat het hoofdprogramma er als volgt uit ziet:
(De werking van het programma wijzigt niet.)

```VBA
Sub oef10_4a()
    Const MIN = 5
    Const MAX = 15

    Dim intA As Integer
        
    intA = CInt(InputBox("Geef getal"))
    If boolGetalTussen(intA, MIN, MAX) Then
        MsgBox ("OK")
    Else
        MsgBox ("NOK")
    End If

End Sub
```


## Oefening 5
Geef volgend programma in en test het:

```VBA
Sub oef10_5()
    Dim strIn1 As String
    Dim strIn2 As String
    Dim strIn3 As String
    Dim strOut As String

    strIn1 = InputBox("Geef minimum, de grenzen zijn 0 en 100")
    strIn2 = InputBox("Geef maximum, de grenzen zijn 0 en 100")
    strIn3 = InputBox("Geef getal, de grenzen zijn " & StrIn1 & " en " & StrIn2)
    strOut = "U gaf in " & vbNewLine
    strOut = strOut & "Minimum: " & strIn1 & vbNewLine
    strOut = strOut & "Maximum: " & strIn2 & vbNewLine
    strOut = strOut & "Getal: " & StrIn3
    
    MsgBox (strOut)

End Sub
```

Pas het programma aan door het bijschrijven van een functie zodat het hoofdprogramma er als volgt uit ziet:
(De werking van het programma wijzigt niet.)

```VBA
Sub oef10_5a()
    Dim strIn1 As String
    Dim strIn2 As String
    Dim strIn3 As String
    Dim strOut As String

    strIn1 = strLeesGetalStringGrenzen("Geef minimum", 0, 100)
    strIn2 = strLeesGetalStringGrenzen("Geef maximum", 0, 100)
    strIn3 = strLeesGetalStringGrenzen("Geef Getal", CInt(strIn1), CInt(strIn2))
    strOut = "U gaf in " & vbNewLine
    strOut = strOut & "Minimum: " & StrIn1 & vbNewLine
    strOut = strOut & "Maximum: " & StrIn2 & vbNewLine
    strOut = strOut & "Getal: " & strIn3

    MsgBox (strOut)
    
End Sub
```

## Oefening 6
Schrijf een functie om een Integer getal in te lezen.

De functie maakt gebruik van de functie InputBox om het getal in te lezen.

Het ingelezen getal wordt eerst in een variabele van het type String geplaatst. Er wordt nagekeken of de invoer effectief een getal is (lus zolang de invoer geen getal is).

Als de invoer een getal is, wordt het getal gekopieerd naar een variabele van het type Double en nagekeken of het getal voldoet aan het criterium Integer: grenzen OK en geen decimalen (lus zolang het getal geen Integer is).

Opmerking:
De programmacode mag gekopieerd worden van vorige gelijkaardige oefeningen.

TIP : een voorbeeld voor de declaratie van deze functie is:

Function intLeesInteger() As Integer

Schrijf een testprogramma waar de functie opgeroepen wordt en de functiewaarde (terugkeerwaarde) in een variabele bewaard wordt. In een volgende lijn wordt de variabele op het scherm geprint (functie MsgBox).
## Oefening 7
Schrijf een booleaanse functie (naam boolIsInteger) met als parameter 1 stringvariabele. De functie geeft True (WAAR) als de stringvariabele een geldig integergetal bevat. In alle andere gevallen geeft deze False (ONWAAR).
## Oefening 8
Pas de functie van oefening  6 aan zodat deze kan gebruikt worden om een Integer getal in te lezen tussen 0 en een maximumwaarde.

De maximumwaarde is een parameter van de functie.

TIP: een voorbeeld voor de declaratie van deze functie is:

Function intLeesIntegerMax(intMax As Integer) As Integer

Schrijf een testprogramma (hoofdprogramma) waar de functie met een for-lus opgeroepen wordt zodat de parameter de waarden 100, 200, 300 en 400 krijgt.

De functiewaarde (terugkeerwaarde) wordt in een variabele bewaard. In een volgende lijn wordt de variabele op het scherm geprint (functie MsgBox).

## Oefening 9
Pas de functie van oefening 8 aan zodat twee parameters meegegeven worden. Deze parameters zijn het minimum en het maximum van het ingegeven getal. De functie mag dus enkel getallen aanvaarden tussen het minimum en het maximum. Het maximum en het minimum worden vermeld in de vraagstelling.

TIP: een voorbeeld voor de declaratie van deze functie is :
```VBA
Function intLeesIntMinMax(intMin As Integer, intMax As Integer) As Integer
```
Schrijf een testprogramma waar de functie met een for-lus opgeroepen wordt zodat de parameter voor het maximum de waarden 100, 200, 300 en 400 krijgt.

De parameter voor het minimum krijgt respectievelijk de waarden -100, -200, -300 en -400.

De functiewaarde (terugkeerwaarde) wordt in een variabele bewaard. In een volgende lijn wordt de variabele op het scherm geprint (functie MsgBox).



## Oefening 10
Breid de functie van oefening 9 uit met een extra parameter. Deze parameter is van het type String en wordt gebruikt om een boodschap mee te geven die moet weergegeven worden bij het oproepen van de functie InputBox. De boodschap wordt in de functie echter nog aangevuld met de grenzen (de parameters minimum en maximum van de functie).

Voorbeeld: De parameter boodschap bevat “Geef een deeltal”, de parameter minimum bevat 5 en de parameter maximum bevat 10. De boodschap die de functie plaatst in de InputBox wordt dan:

“Geef een deeltal, de grenzen zijn 5 en 10”

Maak verder een functie in het hoofdprogramma om 5 elementen in te lezen en te bewaren in een array. Maak hiervoor gebruik van een lus. Bij het aanroepen van de functie moet het nummer van het element in de array bij in de boodschap staan. Hier kan men best eerst een String variabele de juiste inhoud geven (met het gevraagde nummer) en deze String variabele gebruiken als parameter bij het aanroepen van de functie.

Maak een globaal gedeclareerde constante aan. In dit geval 5, het aantal elementen waar we mee gaan werken.

Gebruik deze constante bij de declaratie van de array en bij de for-lus als stopwaarde.

Let op: als we declareren A(5) hebben we een array met 6 elementen, A(0), A(1), ... , A(5).

TIP: een voorbeeld voor de declaratie van deze functie is :

```VBA
Function intLeesIntMinMax(intMin As Integer, intMax As Integer, strBoodschap As String) As Integer
    'Schrijf hier je eigen code
End Function
```


Plaats de 5 ingelezen elementen achter elkaar in een String variabele en geef de String variabele weer met de functie MsgBox.

## Oefening 11
Pas de functie van oefening 10 aan zodat de functie die je geschreven hebt in oefening 5 opgeroepen wordt binnen deze functie.
