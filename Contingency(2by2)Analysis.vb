Option Explicit

Sub cal_sensi_speci()

Dim true_positive, false_positive, false_negative, true_negative As Long
Dim SE
Dim Odds_Ratio As Variant
Dim a, b, c, d As Double

true_positive = Application.WorksheetFunction.CountIfs(Range(Range("A1"), Range("A1").End(xlDown)), "=1", Range(Range("B1"), Range("B1").End(xlDown)), "=1")
false_positive = Application.WorksheetFunction.CountIfs(Range(Range("A1"), Range("A1").End(xlDown)), "=1", Range(Range("B1"), Range("B1").End(xlDown)), "=0")
false_negative = Application.WorksheetFunction.CountIfs(Range(Range("A1"), Range("A1").End(xlDown)), "=0", Range(Range("B1"), Range("B1").End(xlDown)), "=1")
true_negative = Application.WorksheetFunction.CountIfs(Range(Range("A1"), Range("A1").End(xlDown)), "=0", Range(Range("B1"), Range("B1").End(xlDown)), "=0")

Range("D3:M5").Clear

Range("C3") = "Test-Positive"
Range("D3") = true_positive

Range("C4") = "Test-Negative"
Range("D4") = false_negative


Range("E3") = false_positive
Range("E4") = true_negative

'Sensitivty
a = Range("D3").Value + Range("D4").Value
If a <> 0 Then
    Range("H2") = "Sensitivity%"
    Range("H3") = Range("D3").Value / (Range("D3").Value + Range("D4").Value)
    Range("H3").NumberFormat = "0%"
Else:
    Range("H3") = "N/A"
End If

'Specificity
b = Range("E3").Value + Range("E4").Value
If b <> 0 Then
    Range("I2") = "Specifity%"
    Range("I3") = Range("E4").Value / (Range("E4").Value + Range("E3").Value)
    Range("I3").NumberFormat = "0%"
Else:
    Range("I3") = "N/A"
End If

'PPV
c = Range("D3").Value + Range("E3").Value
If a <> 0 Then
    Range("J2") = "Positive Predictive Value%"
    Range("J3") = Range("D3").Value / (Range("D3").Value + Range("E3").Value)
    Range("J3").NumberFormat = "0%"
Else:
    Range("J3") = "N/A"
End If

'NPV
d = Range("D4").Value + Range("E4").Value
If a <> 0 Then
    Range("K2") = "Negative Predictive Value%"
    Range("K3") = Range("E4").Value / (Range("D4").Value + Range("E4").Value)
    Range("K3").NumberFormat = "0%"
Else:
    Range("K3") = "N/A"
End If

'Accuracy
Range("L2") = "Accuracy%"
Range("L3") = (Range("D3").Value + Range("E4").Value) / (Range("D3").Value + Range("D4").Value + Range("E3").Value + Range("E4").Value)
Range("L3").NumberFormat = "0%"

'CCR
If IsNumeric(Range("H3")) And IsNumeric(Range("I3")) Then
    Range("M2") = "CCR%"
    Range("M3") = (Range("H3").Value + Range("I3").Value) / 2
    Range("M3").NumberFormat = "0%"
Else:
    Range("M3") = "N/A"
End If

'PRR
Range("n2") = "PRR"
If Range("e3") = 0 Then
    Range("N3") = "infinity"
Else:
    Range("N3") = Range("D3") * (Range("E3") + Range("E4")) / (Range("D3") + Range("D4")) / Range("E3")
End If
    

Range("F2") = "Odds Ratio(95% CI)"
If Range("E3").Value * Range("D4").Value <> 0 Then
    Odds_Ratio = (Range("D3").Value * Range("E4").Value) / (Range("E3").Value * Range("D4").Value)
    Range("F3") = Odds_Ratio
    Range("F3").NumberFormat = "0.00"
    SE = 1.96 * Sqr(1 / Range("D3").Value + 1 / Range("D4").Value + 1 / Range("E3").Value + 1 / Range("E4").Value)
    Range("F4") = Exp(WorksheetFunction.Ln(Range("f3").Value) - SE)
    Range("F5") = Exp(WorksheetFunction.Ln(Range("f3").Value) + SE)
    Range("F3") = Format(Range("F3").Value, "0.00") & "(" & Format(Range("F4").Value, "0.00") & "-" & Format(Range("F5").Value, "0.00") & ")"
    Range("F4:F5").Clear
Else:
    Range("F3") = "infinity"
End If

Range("G3") = FisherExact(Range("D3:E4"))
Range("G3").NumberFormat = "0.00000"
If Range("G3").Value < 0.001 Then
    Range("G4") = "***"
ElseIf Range("G3").Value < 0.01 Then
    Range("G4") = "**"
ElseIf Range("G3").Value < 0.05 Then
    Range("G4") = "*"
Else
    Range("G4") = ""
End If

Range("C2:M5").HorizontalAlignment = xlCenter
Range("C2:M5").VerticalAlignment = xlCenter

Range(Range("A2"), Range("A2").End(xlDown)).Clear
'Range(Range("B2"), Range("B2").End(xlDown)).Clear

Range("D3:N4").Copy

End Sub

Sub caculate_direct()

Dim SE
Dim Odds_Ratio As Variant
Dim a, b, c, d As Double



Range("C3") = "Test-Positive"
Range("C4") = "Test-Negative"


'Sensitivty
a = Range("D3").Value + Range("D4").Value
If a <> 0 Then
    Range("H2") = "Sensitivity%"
    Range("H3") = Range("D3").Value / (Range("D3").Value + Range("D4").Value)
    Range("H3").NumberFormat = "0%"
Else:
    Range("H3") = "N/A"
End If

'Specificity
b = Range("E3").Value + Range("E4").Value
If b <> 0 Then
    Range("I2") = "Specifity%"
    Range("I3") = Range("E4").Value / (Range("E4").Value + Range("E3").Value)
    Range("I3").NumberFormat = "0%"
Else:
    Range("I3") = "N/A"
End If

'PPV
c = Range("D3").Value + Range("E3").Value
If a <> 0 Then
    Range("J2") = "Positive Predictive Value%"
    Range("J3") = Range("D3").Value / (Range("D3").Value + Range("E3").Value)
    Range("J3").NumberFormat = "0%"
Else:
    Range("J3") = "N/A"
End If

'NPV
d = Range("D4").Value + Range("E4").Value
If a <> 0 Then
    Range("K2") = "Negative Predictive Value%"
    Range("K3") = Range("E4").Value / (Range("D4").Value + Range("E4").Value)
    Range("K3").NumberFormat = "0%"
Else:
    Range("K3") = "N/A"
End If

'Accuracy
Range("L2") = "Accuracy%"
Range("L3") = (Range("D3").Value + Range("E4").Value) / (Range("D3").Value + Range("D4").Value + Range("E3").Value + Range("E4").Value)
Range("L3").NumberFormat = "0%"

'CCR
If IsNumeric(Range("H3")) And IsNumeric(Range("I3")) Then
    Range("M2") = "CCR%"
    Range("M3") = (Range("H3").Value + Range("I3").Value) / 2
    Range("M3").NumberFormat = "0%"
Else:
    Range("M3") = "N/A"
End If

Range("F2") = "Odds Ratio(95% CI)"
If Range("E3").Value * Range("D4").Value <> 0 Then
    Odds_Ratio = (Range("D3").Value * Range("E4").Value) / (Range("E3").Value * Range("D4").Value)
    Range("F3") = Odds_Ratio
    Range("F3").NumberFormat = "0.00"
    SE = 1.96 * Sqr(1 / Range("D3").Value + 1 / Range("D4").Value + 1 / Range("E3").Value + 1 / Range("E4").Value)
    Range("F4") = Exp(WorksheetFunction.Ln(Range("f3").Value) - SE)
    Range("F5") = Exp(WorksheetFunction.Ln(Range("f3").Value) + SE)
    Range("F3") = Format(Range("F3").Value, "0.00") & "(" & Format(Range("F4").Value, "0.00") & "-" & Format(Range("F5").Value, "0.00") & ")"
    Range("F4:F5").Clear
Else:
    Range("F3") = "infinity"
End If

Range("G3") = FisherExact(Range("D3:E4"))
Range("G3").NumberFormat = "0.00000"
If Range("G3").Value < 0.001 Then
    Range("G4") = "***"
ElseIf Range("G3").Value < 0.01 Then
    Range("G4") = "**"
ElseIf Range("G3").Value < 0.05 Then
    Range("G4") = "*"
Else
    Range("G4") = ""
End If


Range("C2:M5").HorizontalAlignment = xlCenter
Range("C2:M5").VerticalAlignment = xlCenter

End Sub

Function FisherExact(r As Range) As Double
    Sheets(1).Range("D3:E4").Copy
    Sheets("p-value").Range("B11:C12").PasteSpecial
    FisherExact = Sheets("p-value").Range("C7").Value
End Function
