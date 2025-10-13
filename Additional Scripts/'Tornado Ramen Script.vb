'Tornado Ramen Script 
Dim GlucoseReading as Double = p.ExtE 
Dim GlucoseTarget as Double = 1 '[g/L]
Dim GlucosePumpSP as Double = 10 '[mL/H]

If GlucoseReading < GlucoseTarget Then 
    p.pumpDActive = True 
    p.FDSP = GlucosePumpSP '[mL/H]
ElseIf GlucoseReading > GlucoseTarget Then 
   p.pumpDActive = False
End If 