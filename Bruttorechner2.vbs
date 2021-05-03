' Eingabe
netto = InputBox("Geben Sie den Netto-Betrag ein!","Eingabe Netto")
IF NOT IsNumeric(netto) THEN 'wenn Buchstaben eingegeben wurden 
   MsgBox "Keine Zahl eingegeben!",vbCritical,"Fehler"
   WScript.quit()  'Programmende
END IF
IF isEmpty(netto) THEN  'wenn Abbrechen, dann ist netto leer
   MsgBox "Abbrechen geklickt",,"Abbruch"
   WScript.quit()  ' Programmende
END IF
' Verarbeitung (Berechnung)
mwst = 0.19 * netto
brutto = netto + mwst
'Ausgabe
ergebnis = "Netto:" & vbTab & vbTab & FormatCurrency(netto) & _
           vbNewline & _
           "+ 19% MwSt:" & vbTab & FormatCurrency(mwst) & _
           vbNewline & _
           "Gesamt: " & vbTab & vbTab & FormatCurrency(brutto)


MsgBox ergebnis,,"Ergebnis"