# Basisoefeningen les 11
Bij deze oefenreeks is het belangrijk dat je bij de parameters van de subroutines steeds de vermelding byval of byref wordt bijgevoegd
## Oefening 1
Schrijf een subroutine met 2 integer-getallen als parameter. Tijdens het uitvoeren van deze routine wordt de inhoud van de 2 parameters gewisseld.

Test de subroutine met volgend hoofdprogramma:

```VBA
Sub Oef11_1()
    Dim intA As Integer
    Dim intB As Integer
    
    intA = 10
    intB = 20
    
    MsgBox ("IntA voor Swap: " & CStr(intA))
    MsgBox ("Intb voor Swap: " & CStr(intB))
    
    Call WisselInt(intA, intB)
    
    MsgBox ("IntA na Swap: " & CStr(intA))
    MsgBox ("Intb na Swap: " & CStr(intB))

End Sub
```


## Oefening 2
Schrijf een subroutine om een stringarray weer te geven in de tweede kolom van het excel werkblad.

Test het programma met volgend hoofdprogramma

```VBA
Sub oef11_2()
    Const AANTAL = 5
    
    Dim strGegevens(AANTAL) As String
    Dim intI As Integer
    
    For intI = 1 To 5’
        strGegevens(intI) = "Testgegevens" & CStr(intI)
    Next intI

    Call Array_to_excel(strGegevens(), AANTAL)  

End Sub

```

## Oefening 3
Pas de subroutine van oefening 2 aan zodat de gegevens vanaf de 2de rij worden weergegeven en op de eerste rij een boodschap komt die wordt meegegeven als parameter. De kolom waar de gegevens worden weergegeven is eveneens een parameter. Pas het hoofdprogramma aan zodat de nieuwe subroutine kan gebruikt worden.

## Oefening 4
Schrijf een subroutine om een stringarray in te lezen. Het array en het aantal in te lezen elementen zijn de parameters. Pas het hoofdprogramma van oefening 3 aan zodat de nieuwe functie kan gebruikt worden. Het weergeven van het array op het Excel werkblad gebeurt met de subroutine die je geschreven hebt in oefening 3

## Oefening 5
Schrijf een functie om een stringarray in te lezen. Het aantal elementen dat wordt ingelezen is variabel. Er wordt gestopt met een leeg veld in te geven. Wanneer het aantal elementen van het array (gedefinieerd in het hoofprogramma) overschreden is, moet het programma stoppen. Het aantal elementen dat gelezen is, is de returnwaarde van de functie, het array komt terug via een parameter.

Schrijf het geschikte hoofdprogramma om dit te testen

## Oefening 6
Schrijf een subroutine om alle elementen van een integer array met 2 te vermenigvuldigen. De subroutine heeft 3 parameters:

- Het invoerarray
- Het uitvoerarray
- Het aantal elementen

Kies het juiste variabeletype voor de array’s zodat geen error kan ontstaan

## Oefening 7
Schrijf een functie van het type boolean. De functie heeft tot doel om de elementen van een integerarray (parameter) te vermenigvuldigen met 2. De functie heeft 2 parameters:

- Array
- Het aantal elementen

Er mag echter geen error (overflow) ontstaan wanneer een bewerking niet kan uitgevoerd worden. Indien het array niet de juiste elementen bevat moet het vermenigvuldigen stoppen en moet de functie ONWAAR retourneren. Wanneer het uitvoeren van de functie perfect verlopen is moet de functie WAAR retourneren



