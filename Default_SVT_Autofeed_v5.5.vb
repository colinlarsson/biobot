'Autofeed for Scivario - Generated Configuration
'Version 5.4 - Generated: 2025-11-04
'User: Colin
'Experiment: Default
'Reactor: SVT

'=== PUMP CONFIGURATION ===
'Feed A (Major Feed): Pump D - Gravimetric - FCal: 17.2
'Feed B: Pump F - Volumetric - FCal: 105
'Glucose: Pump E - Gravimetric - FCal: 28

'Feed Variables
Dim MajorFeed_Fed as Double '[g]
Dim glucoseFeed_Fed as Double '[g]

'BalanceState Function
'Case 0, Case 1, Case 7, Case 8, Case 9, Case 12, Case 14
Dim State As Integer = p.BalanceAState()

Dim Scale_Weight(6) as Double 
Scale_Weight(1) = p.MAPV 'time zero bottle weight 
Scale_Weight(2) = p.MAPV 'interval starting bottleweight 
Scale_Weight(3) = p.MAPV 'current weight feed A
Scale_Weight(4) = p.MAPV 'glucose interval starting bottleweight
Scale_Weight(5) = p.MAPV 'current weight glucose

'Process Day Calculation
Dim procDay As Integer
'Caluation for Process Day, rolls over 8 hours before inoculation time
Dim calculatedProcDay As Integer
calculatedProcDay = int((p.inoculationTime_H + 8) / 24)

'Process day logic, 8 hours before inoculation time, the day ticks over.
If calculatedProcDay < 0 Then
    procDay = 0
Else
    procDay = calculatedProcDay
End If

'Configuration Parameters
Dim sampleDelay as Double = 4
'Feed Start Time in hours, initialized to 0
Dim Feed_StartTime as Double = 0
'Lead time for sample prompt in hours, not used right now
Dim samplePromptLead_H As Double = 4

'Feed Schedule Arrays
Dim Feed_Time(15) as Double
Feed_Time(1) = 24
Feed_Time(2) = 48
Feed_Time(3) = 72
Feed_Time(4) = 96
Feed_Time(5) = 120
Feed_Time(6) = 144
Feed_Time(7) = 168
Feed_Time(8) = 192
Feed_Time(9) = 216
Feed_Time(10) = 240
Feed_Time(11) = 264
Feed_Time(12) = 288
Feed_Time(13) = 312
Feed_Time(14) = 336

Dim Major_Feed_TargetWeight as Double
Dim Major_Feed_TargetWeight_Value(15) as Double
Major_Feed_TargetWeight_Value(1) = 0
Major_Feed_TargetWeight_Value(2) = 45
Major_Feed_TargetWeight_Value(3) = 0
Major_Feed_TargetWeight_Value(4) = 90
Major_Feed_TargetWeight_Value(5) = 0
Major_Feed_TargetWeight_Value(6) = 120
Major_Feed_TargetWeight_Value(7) = 0
Major_Feed_TargetWeight_Value(8) = 150
Major_Feed_TargetWeight_Value(9) = 0
Major_Feed_TargetWeight_Value(10) = 150
Major_Feed_TargetWeight_Value(11) = 0
Major_Feed_TargetWeight_Value(12) = 0
Major_Feed_TargetWeight_Value(13) = 0
Major_Feed_TargetWeight_Value(14) = 0

Dim FeedB_VolumeTarget as Double = 0
Dim FeedB_VolumeTarget_Value(15) as Double
FeedB_VolumeTarget_Value(1) = 0
FeedB_VolumeTarget_Value(2) = 1.5
FeedB_VolumeTarget_Value(3) = 0
FeedB_VolumeTarget_Value(4) = 3
FeedB_VolumeTarget_Value(5) = 0
FeedB_VolumeTarget_Value(6) = 4.5
FeedB_VolumeTarget_Value(7) = 0
FeedB_VolumeTarget_Value(8) = 6
FeedB_VolumeTarget_Value(9) = 0
FeedB_VolumeTarget_Value(10) = 7.5
FeedB_VolumeTarget_Value(11) = 0
FeedB_VolumeTarget_Value(12) = 0
FeedB_VolumeTarget_Value(13) = 0
FeedB_VolumeTarget_Value(14) = 0

'Pump Calibrations
Dim FeedA_FCal as Double = 17.2
Dim FeedB_FCal as Double = 105
Dim Glucose_FCal as Double = 28

' How many scheduled events are configured (easy to change)
Dim MaxFeeds As Integer = 14

'Feed A SetPoints
Dim Major_Feed_SP as Double = 250
Dim Major_Feed_Percent_Value as Double = 0.25 'reduce pump flow rate to 25% of SP 
Dim Major_Feed_SP_percent as Double = Major_Feed_SP * Major_Feed_Percent_Value 'stores the new reduced pump flow sp
Dim cut_pumpValue_Percent as Double = 0.75 'slow down when reached 75% of target value 
Dim Major_Feed_Counter as Double = 0

'Glucose Feed SetPoints
Dim Glucose_Feed_SP as Double = 160
Dim Glucose_Feed_Percent_Value as Double = 0.25
Dim Glucose_Feed_SP_percent as Double = Glucose_Feed_SP * Glucose_Feed_Percent_Value
Dim Glucose_Cut_pumpValue_Percent as Double = 0.75

'Feed B Parameters
Dim FeedB_FlowRate as Double = 35
Dim FeedB_Counter as Double = 0
Dim FeedB_Totalizer as Double = 0

'Initialize State Array
If s is Nothing then
    Dim a(22) as Double
    a(1) = 0  'time zero bottle weight
    a(2) = 0  'interval starting bottleweight
    a(3) = 0  'current weight
    a(4) = Major_Feed_TargetWeight
    a(5) = Major_Feed_Counter
    a(6) = Feed_StartTime
    a(7) = FeedB_Counter
    a(8) = FeedB_Totalizer
    a(9) = FeedB_VolumeTarget
    a(10) = 0 'Glucose interval starting bottleweight
    a(11) = 0 'current weight for glucose
    a(12) = 0 'user input glucose
    a(13) = 0 'Feed A totalizer baseline
    a(14) = 0 'Glucose totalizer baseline
    a(15) = 0 'Feed B totalizer baseline
    a(16) = 0 'Glucose and Sample Confirmation Flag
    s = a
End If

If P isNot Nothing Then
    With P 

    'interval starting weight minus current weight
    MajorFeed_Fed = (s(2) - s(3)) 
    'glucose feed starting weight minus current weight
    glucoseFeed_Fed = (s(10) - s(11)) 

    '.intA =  s(12) 'requested glucose input
    '.intB = sample confirmation, 0 = not confirmed, 1 = confirmed
    .intC = procDay 'show process day  
    .intD = s(5) 'Feed Counter
    .intE = s(4) 'Feed A Target
    '.intF = s(2) - s(3) 'Feed A Fed Update: Defined later in feed A block
    .intG = s(12) 'glucose target display
    .intH = s(10) - s(11) 'glucose fed
    .intI = s(9) + s(8) 'feed b volume target
    .IntJ = .VFPV - s(8) 'Track Feed B Fed, not true value
    .intK = s(2) 'static interval feed
    .intL = s(3) 'dynamic scale reading during phase 9
    '.intT = s(6) - .inoculationTime_H 'time to feed in hours. changing, won't be set until after phase 3 runs
    .intU = .phase 'phase
    .intV = .runtime_H - .phaseStart_H 'phase time

    'Calculate time remaining until next feed
    Dim timeRemaining As Double

    If s(5) = 0 Then 'sets intT to use the first feed time if the feed counter s(5) hasn't turned over
        timeRemaining = Feed_Time(1) - .inoculationTime_H 'time to first feed in hours.
    Else 
        timeRemaining = s(6) - .inoculationTime_H 's(6) will update after phase 3
    End If

    If timeRemaining < 0 Then 'Prevents time remaining intT from being negative
        .intT = 0
    Else
        .intT = timeRemaining
    End If
    
Static lastState As Integer

Select Case .phase

    Case 0 'Workflow Start 
        .logmessage("Autofeed Scivario Test Loaded")
        .phase = .phase + 1

'==========================================================================================
' PHASE 1 - Start Inoculation Timer
'==========================================================================================

            Case 1

                If .InoculationTime_H > 0 And (.runtime_H - .phaseStart_H) > 0.1/60 Then
                    .phase = .phase + 1
                End If

'==========================================================================================
' PHASE 2 - Counter Management and totalizer reset
'==========================================================================================

            Case 2

                .pumpDActive = False
                .pumpFActive = False
                .pumpEActive = False
                
                'Reset the intA/intB reset flag to 0, should made them user adjustable
                
                If .OfflineM = 0 Then 'Update Counters Manually, M: 1 enter reset state, N: number of steps to skip, O: step up 1. M will intercept after every feed
                'not currently functional, will be for pause function

                    If .InoculationTime_H > Feed_Time(1) And (.runtime_H - .phaseStart_H) > 0.1/60 Then 'waits for first feed time + 6s delay;
                        s(5) = s(5) + 1
                        s(7) = s(7) + 1
                        .phase = .phase + 1
                    ElseIf s(5) > MaxFeeds Then 'checks if all feeds are done
                        .pumpDActive = False
                        .pumpFActive = False
                        .pumpEActive = False
                        .logmessage("All Feeds Delivered: Major=" & CInt(s(5)) & "/" & MaxFeeds & ", Feed B=" & CInt(s(7)) & "/" & MaxFeeds & ". Stopping (Phase 16).")
                        .phase = 16
                    End If
                End If
 
'==========================================================================================
' PHASE 3 - Assign New Values for Interval, looks for sample confirmation
'==========================================================================================

            Case 3

                'Assign feed parameters based on counter
                If s(5) >= 1 And s(5) <= 14 Then
                    Feed_StartTime = Feed_Time(s(5))
                    Major_Feed_TargetWeight = Major_Feed_TargetWeight_Value(s(5))
                End If
                
                If s(7) >= 1 And s(7) <= 14 Then
                    FeedB_Totalizer = .VFPV
                    FeedB_VolumeTarget = FeedB_VolumeTarget_Value(s(7))
                End If
                
                'store feed parameters in state array
                s(4) = Major_Feed_TargetWeight
                s(6) = Feed_StartTime
                s(8) = FeedB_Totalizer
                s(9) = FeedB_VolumeTarget
                
                'if sample not confirmed before feed time, will sit and wait in next phase. this logmessage runs a day early
                If .intB <> procDay And .inoculationTime_H < (s(6) + sampleDelay) And ((.runtime_H - .PhaseStart_H) > 0.1 / 60) Then 'Waiting for intB to confirm sample
                    .logwarning("██ - Phase: 3 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ". - ████ - Enter Glucose Target in internal A, then confirm sample in internal B fields.")
                    .intT = s(6) - .inoculationTime_H 'time to feed in hours. 
                    .phase = .phase + 1
                ElseIf .intB <> procDay And .inoculationTime_H > (s(6) + sampleDelay) Then 'if sample is not confirmed before phase 4, will skip next phase
                    .logwarning("██ - Phase: 3 - Day " & procDay & " - Unit #: " & .unit & ". - ██ - Sample not confirmed, feeding late, no glucose entered")
                    .intT = s(6) - .inoculationTime_H 'time to feed in hours. 
                    .phase = .phase + 2
                ElseIf .intB = procDay And .inoculationTime_H > s(6) 'if sample is confirmed before phase 4, this message will display, unlikely to happen
                    .intT = s(6) - .inoculationTime_H 'time to feed in hours. 
                    .logmessage("██ - Phase: 3 - Day " & procDay & " - Unit #: " & .unit & ". - ██ - Sampling Confirmed: Next Feeding Event in " & formatNumber(.intT, 2) & " Hrs")
                    .logmessage("██ - Phase: 3 - Day " & procDay & " - Unit #: " & .unit & ". - ██ - Feed Targets: " & s(9) & " [mL] FB. " & S(4) & " [g] FA and " & .intA & " [g] Glucose.")
                    .phase = .phase + 2
                End If
 
              
'==========================================================================================
' PHASE 4 - Waiting for Sample Confirmation
'==========================================================================================
'sits here after first feed, waiting for sample confirmation


            Case 4

                s(12) = .intA 'update requested glucose

                'If sample not confirmed before delay time is passed, will give warning and go to feedB
                If .intB <> procDay And .inoculationTime_H > (s(6) + sampleDelay) Then 
                    .logwarning("Phase 4: Sample not verified, feeding late, no glucose entered")
                    .phase = .phase + 1
                'If sample confirmed, will check feed time and proceed
                ElseIf .intB = procDay and .inoculationTime_H > s(6) Then 
                    .logmessage("██ - Phase: 4 - Day " & procDay & " - Unit #: " & .unit & ". - ██ - Sampling Confirmed: Next Feeding Event in " & formatNumber(.intT, 2) & " Hrs")
                    .logmessage("██ - Phase: 4 - Day " & procDay & " - Unit #: " & .unit & ". - ██ - Feed Targets: " & s(9) & " [mL] FB. " & S(4) & " [g] FA and " & .intA & " [g] Glucose.")
                    .phase = .phase + 1
                End If

'==========================================================================================
' PHASE 5 - Determines if Feed B is requested
'==========================================================================================

            Case 5

                'Skips to phase 7 if no feed B is requested
                If ((.runtime_H - .PhaseStart_H) > 0.1/60) And .inoculationTime_H > s(6) And s(9) <> 0 Then '6 second delay to ensure values get stored to array. also accounts for hiccups in DWC, checks for fb entry
                    .logmessage("Phase: 5.  Day: " & procDay & ".  Unit #: " & .unit & ".  ---Feed B Starting: " & s(9) & " [mL] to be fed.")
                    .phase = .phase + 1
                ElseIf ((.runtime_H - .PhaseStart_H) > 0.1/60) And .inoculationTime_H > s(6) And s(9) <= 0 Then 
                    .logmessage("Phase: 5.  Day: " & procDay & ".  Unit #: " & .unit & ".  ---No Feed B To be fed, skipping.")
                    .phase = .phase + 2 'goes to phase 7
                End If

'==========================================================================================
' PHASE 6 - Feed B Start
'==========================================================================================                

            Case 6

                '*************PHASE VARIABLES*************

                's(7) = FeedB Counter, redundand with s(5)
                's(8) = FeedB_Totalizer: starting totalizer snapshot from phase 3
                's(9) = FeedB_VolumeTarget: volume chosen in phase 3

                '*************PHASE VARIABLES*************

                If .inoculationTime_H > s(6) Then
                    .pumpFActive = True 'turn feed b on
                End If

                If s(7) <= MaxFeeds Then
                    'if totalizer is less than target set in phase 3, keep feeding at specified flowrate
                    If .VFPV < s(8) + s(9) Then
                        .FFSP = FeedB_FlowRate
                    'when totalizer exceeds target, turn feed B off and move to next phase
                    Else
                        .logmessage("██ - Phase: 6 - Day " & procDay & " - Unit #: " & .unit & ". - ██ - Feed B Finished: " & s(9) & " [mL] Fed")
                        .phase = .phase + 1
                    End If
                ElseIf s(7) > MaxFeeds And (.runtime_H - .phaseStart_H) > 0.1/60 Then
                    .phase = 16
                End If

                
'==========================================================================================
' PHASE 7 - Feed B Off, Next Step
'==========================================================================================

            Case 7

                .pumpFActive = False

                'takes snapshot of scale weight after feed B, for feed A
                s(1) = Scale_Weight(1)
                
                'max feed check - not necessary
                If .inoculationTime_H > Feed_Time(14) Then
                    .phase = 2
                Else
                    .phase = .phase + 1
                End If
 
                
'==========================================================================================
' PHASE 8 - Prepare for feed A
'==========================================================================================

            Case 8

                '*************PHASE VARIABLES*************

                'Record Feed A totalizer baseline at start of interval, should not update in phase 9
                s(13) = .VDPV

                'Record Scale weight and set s(2) to starting weight for feed A
                s(2) = Scale_Weight(2) 
                'Set dynamic weight to s(2) just in-case it were negative, will begin updating in phase 9
                s(3) = s(2) 

                's(6) = Feed_StartTime from phase 3
                's(4) = Major_Feed_TargetWeight from phase 3

                '*************PHASE VARIABLES END*************
                
                'Waits for respective feed time + 6s delay; then moves to feed A phase
                If .inoculationTime_H > s(6) And s(4) <> 0 And (.runtime_H - .phaseStart_H) > 0.1/60 Then
                    .phase = .phase + 1
                'If no feed A, skips to phase 10
                ElseIf s(4) = 0 Then
                    .logmessage("Phase: 8.  Day: " & procDay & ".  Unit #: " & .unit & ".  No Feed A entered, skipping Feed A phase.")
                    .phase = .phase + 2
                End If
                

'==========================================================================================
' PHASE 9 - Feed A Runs
'==========================================================================================

            Case 9

                '*************PHASE VARIABLES*************

                'Dynamic feed A totalizer calculation, updates live during phase 9
                Dim FeedA_Fed_Totalizer As Double
                FeedA_Fed_Totalizer = .VDPV - s(13) '[g]
                
                s(3) = Scale_Weight(3)

                '.FDSP Dasware pump D flowrate setpoint; [mL/H]
                'Major_Feed_SP = 250 '[mL/H] - starting pump flow setpoint, set in variable section
                'Major_Feed_SP_percent = Major_Feed_SP * 0.25  'stores the new reduced pump flow sp, set in variable section  
                'cut_pumpValue_Percent = 0.75 'slow down when reached 75% of target value , set in variable section
                's(2) = interval starting bottleweight from phase 8
                's(3) = current weight feed A from scale
                's(4) = Major_Feed_TargetWeight from phase 3, updates each day
                'MajorFeed_Fed = (s(2) - s(3)) '[g] set in variable section
                's(13) = Feed A totalizer baseline from phase 8

                '*************PHASE VARIABLES*************

                'intF is display for Feed A fed, updates live during phase 9, phase 8 snapshot weight minus current weight  
                .intF = s(2) - s(3) '[g] same as MajorFeed_Fed

                'feed A active true
                .pumpDActive = True
                
                'Speed Control Logic
                'if majorfeedfed exceeds 75% of target weight, reduce flowrate to 25% of starting setpoint
                If Abs(MajorFeed_Fed) > (s(4) * cut_pumpValue_Percent) Then
                    .FDSP = Major_Feed_SP_percent 'reduced pump flow setpoint 
                Else
                    .FDSP = Major_Feed_SP 'starting pump flow rate
                End If
                
                'Feed Completion Logic
                'If majorfeedfed exceeds target weight, stop feed
                If Abs(MajorFeed_Fed) > s(4) And (.runtime_H - .phaseStart_H) > 0.1/60 Then '(s(1) - s(2)) > s(12)
                    .pumpDActive = False
                    .intF = s(2) - s(3) 'Feed A Fed by scale, probably not necessary since updated live
                    .logmessage("██ - Phase: 9 - Day " & procDay & " - Unit #: " & .unit & ". - ██ Feed A Fed:  " & formatNumber(MajorFeed_Fed, 2) & " [g]")
                    .phase = .phase + 1
                'If dynamic totalizer is 110% above target weight, stop feed
                ElseIf (.runtime_H - .phaseStart_H) > 0.1/60 And FeedA_Fed_Totalizer > (1.1 * s(4)) Then 'if dynamic totalizer is above 110% of target weight then stop
                    .pumpDActive = False
                    .intF = FeedA_Fed_Totalizer 'record dynamic totalizer value
                    .logalarm("██ - Phase: 9 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ".██ Feed A Totalizer: " & formatNumber(FeedA_Fed_Totalizer, 2) & " [g] is above 110% of target weight: " & s(4) & " [g]. Stopping Feed A.")
                    .phase = .phase + 1
                End If
                

                'if need to reset loop, then set offline C = 2. when in phase 2, set offline C = 0 to assign phase, then once in phase 4 assign offline C = 1   
                If .OfflineO = 2 Then
                    .phase = 2
                End If

            
'==========================================================================================
' PHASE 10 - Glucose Decision Tree
'==========================================================================================

            Case 10

                s(10) = Scale_Weight(4) 'sets starting glucose interval scale weight
                s(11) = s(10) 'Initialize end weight to prevent negative display if feed is skipped
                s(12) = .intA 'requested glucose
                
                'intB <> procDay means sample not confirmed
                'intB = procDay means sample confirmed
                '(s(6) + sampleDelay) = feed time + delay time

                  If .intB <> procDay AndAlso .inoculationTime_H > s(6) Then 'skips glucose if sample not confirmed
                      .phase = .phase + 1 'moves to waiting for sample confirmation again
                  ElseIf .intB = procDay Then
                      .phase = .phase + 2
                  End If

'==========================================================================================
' PHASE 11 - Wait for sample Confirmation
'==========================================================================================

            Case 11

                'intB <> procDay means sample not confirmed
                'intB = procDay means sample confirmed
                '(s(6) + sampleDelay) = feed time + delay time
                s(12) = .intA 'requested glucose
                
                If .intB <> procDay and .inoculationTime_H > (s(6) + sampleDelay) Then 'skips glucose if sample not confirmed
                    .logwarning("██ - Phase: 11 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ". - ████ - Skipping glucose feed, sample not verified")
                    s(11) = s(10) 'show glucose fed as 0
                    .phase = .phase + 4
                ElseIf .intB = procDay Then 'if not entered until after feed
                    .phase = .phase + 1
                End If

'==========================================================================================
' PHASE 12 - Record Glucose Phase Bottleweight Baseline and Totalizer Baseline
'==========================================================================================

            Case 12

                s(10) = Scale_Weight(4) 'Static glucose interval starting scale weight
                s(11) = s(10) 'Initialize end weight to prevent negative display
                s(12) = .intA 'update requested glucose
                s(14) = .VEPV 'Record Glucose totalizer baseline at start of interval 
                
                If .inoculationTime_H > s(6) And (.runtime_H - .phaseStart_H) > 0.5/60 Then 'waits for respective feed time + 30s delay; 
                    .phase = .phase + 1
                End If

'==========================================================================================
' PHASE 13 - Prepare for Glucose
'==========================================================================================

            Case 13

                '*************PHASE VARIABLES*************

                'Record Glucose totalizer baseline at start of interval, should not update in phase 14
                s(14) = .VEPV '[mL]

                'Record Scale weight and set s(10) to starting weight for glucose
                s(10) = Scale_Weight(4) '[g]

                'Set dynamic weight to s(10) just in-case it were negative, will begin updating in phase 14
                s(11) = s(10) '[g]

                's(12) = user input glucose from phase 4 or 11

                '*************PHASE VARIABLES END*************
                
                If s(12) = 0 AndAlso (.runtime_H - .phaseStart_H) > 0.5/60 Then 'no glucose to be added, skips glucose feed
                    .logmessage("Phase: 14.  Day: " & procDay & ".  Unit #: " & .unit & ".  ---Sampling Confirmed: " & s(12) & " [g] Glucose to be fed.")
                    .phase = .phase + 2
                ElseIf s(12) <> 0 AndAlso (.runtime_H - .phaseStart_H) > 0.5/60 Then 
                    .logmessage("Phase: 14.  Day: " & procDay & ".  Unit #: " & .unit & ".  ---Sampling Confirmed: " & s(12) & " [g] Glucose to be fed.")
                    .phase = .phase + 1
                End If

'==========================================================================================
' PHASE 14 - Glucose Feed
'==========================================================================================               
                
            Case 14

                '*************PHASE VARIABLES*************

                'Dynamic glucose totalizer calculation, updates live during phase 14
                Dim Glucose_Fed_Totalizer As Double '[g]
                Glucose_Fed_Totalizer = .VEPV - s(14) '[g]
                
                'Dynamic glucose fed calculation, updates live during phase 14
                s(11) = Scale_Weight(5) '[g]

                'intH is display for glucose fed, updates live during phase 14, phase 13 snapshot weight minus current weight'
                .intH = s(10) - s(11) '[g] same as glucoseFeed_Fed

                's(12) = user input glucose
                'glucose_feed_sp = 150 '[mL/H] - starting pump flow setpoint, set in variable section
                'Glucose_Feed_SP_percent = glucose_feed_sp * 0.25  'stores the new reduced pump flow sp, set in variable section  
                'Glucose_Cut_pumpValue_Percent = 0.75 'slow down
                's(10) = starting glucose weight
                's(11) = current glucose weight
                's(12) = user input glucose target weight
                'glucoseFeed_Fed = (s(10) - s(11)) [g] 'calculated in variable section
                's(14) = Glucose totalizer baseline from phase 13

                '*************PHASE VARIABLES*************

                .pumpEActive = True
                
                'Speed Control Logic
                'if glucosefeedfed exceeds 75% of target weight, reduce flowrate to 25% of starting setpoint
                If Abs(glucoseFeed_Fed) > (s(12) * Glucose_Cut_pumpValue_Percent) Then
                    .FESP = Glucose_Feed_SP_percent
                Else
                    .FESP = Glucose_Feed_SP
                End If
                
                If Abs(glucoseFeed_Fed) > s(12) And (.runtime_H - .phaseStart_H) > 0.1/60 Then
                    .logmessage("Phase: 14  Day: " & procDay & ".  Unit #: " & .unit & ".  ---Feed Counter: [" & .intD &"]")
                    .phase = .phase + 1
                ElseIf (.runtime_H - .phaseStart_H) > 0.1/60 And Glucose_Fed_Totalizer > (1.1 * s(12)) Then
                    .pumpEActive = False
                    .logwarning("██ - Phase: 14 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ". - ████ - Totalizer: " & formatNumber(.VEPV, 2) & " [g] is above 110% of target weight: " & s(12) & " [g]. Stopping Feed A.")
                    .phase = .phase + 1
                End If

'==========================================================================================
' PHASE 15 - Wrap Up
'==========================================================================================

            Case 15

                .pumpEActive = False
                
                
                If .runtime_H - .phaseStart_H > 0.1/60 Then
                    .logwarning("██ - Phase: 15 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ". - ████ - Total Feeds: - ██ - Feed A: " & formatNumber(MajorFeed_Fed, 2) & " [g]. Feed B: " & s(9) & " [mL]. Glucose:  " & formatNumber(glucoseFeed_Fed, 2) & "[g] - ██")
                    .phase = 2
                End If
                
            Case 16
                .pumpDActive = False
                .pumpFActive = False
                .pumpEActive = False
                
        End Select
    End With
End If

'Debug logs (disable when not debugging)
'p.LogMessage("Phase: " & p.phase)
'p.LogMessage("Inoc Time [h]: " & p.InoculationTime_H)
'p.LogMessage("MajorFeed Fed: " & MajorFeed_Fed)
'p.LogMessage("Array s(2) - STATIC INTERVAL: " & s(2))
'p.LogMessage("Array s(3) - DYNAMIC INTERVAL: " & s(3))
'p.LogMessage("Array s(4) - TARGET: " & s(4))
'p.LogMessage("Array s(5) - COUNTER: " & s(5))
