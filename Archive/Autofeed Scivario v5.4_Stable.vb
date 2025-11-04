'Autofeed for Scivario
'Version 5.0 Last Update 2025-08-09

'E = Glucose gravimetric f.cal = 24
'B = Base not controlled
'D = Feed A gravimetric f.cal = 15.92
'F = Feed B volumetric f.cal = 105

'pumpDActive
'VDPV
'FDSP
'fdcal

'Feed A 
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

'Process Day
Dim procDay As Integer 'Day used for logmessages

'Caluation for Process Day, rolls over 8 hours before inoculation time
Dim calculatedProcDay As Integer
calculatedProcDay = int((p.inoculationTime_H + 8) / 24)

'Process day logic, 8 hours before inoculation time, the day ticks over.
If calculatedProcDay < 0 Then
    procDay = 0
Else
    procDay = calculatedProcDay
End If

Dim sampleDelay as Double

Dim Feed_StartTime as Double = 0 'Feed Start Time in hours, initialized to 0

Dim samplePromptLead_H As Double = 4 'Lead time for sample prompt in hours

'All Feed Times in hours
Dim Feed_Time(15) as Double 

'Major Feed active target weight 
Dim Major_Feed_TargetWeight as Double 

'Major Feed - pump C
Dim Major_Feed_TargetWeight_Value(15) as Double 

'Feed B active target volume
Dim FeedB_VolumeTarget as Double = 0

'Feed B Target Volumes - pump A
Dim FeedB_VolumeTarget_Value(15) as Double 

'Feed A fcal (pump D) = 15.92
'Glucose fcal (pump E) = 24
'feed B fcal (pump F) = 105
Dim pumpD_FCal as Double
Dim pumpF_FCal as Double
Dim pumpE_FCal as Double

' How many scheduled events are configured (easy to change)
Dim MaxFeeds As Integer

pumpE_FCal = 24
pumpD_FCal = 15.92
pumpF_FCal = 105

MaxFeeds = 14

sampleDelay = 5/60

Feed_Time(1) = 5/60
Feed_Time(2) = 10/60
Feed_Time(3) = 15/60
Feed_Time(4) = 1
Feed_Time(5) = 5
Feed_Time(6) = 18
Feed_Time(7) = 24
Feed_Time(8) = 252
Feed_Time(9) = 500
Feed_Time(10) = 501
Feed_Time(11) = 502
Feed_Time(12) = 503
Feed_Time(13) = 504
Feed_Time(14) = 505

Major_Feed_TargetWeight_Value(1) = 20
Major_Feed_TargetWeight_Value(2) = 25
Major_Feed_TargetWeight_Value(3) = 10
Major_Feed_TargetWeight_Value(4) = 22
Major_Feed_TargetWeight_Value(5) = 96.3
Major_Feed_TargetWeight_Value(6) = 96.3
Major_Feed_TargetWeight_Value(7) = 96.3
Major_Feed_TargetWeight_Value(8) = 96.3
Major_Feed_TargetWeight_Value(9) = 96.3
Major_Feed_TargetWeight_Value(10) = 10
Major_Feed_TargetWeight_Value(11) = 11
Major_Feed_TargetWeight_Value(12) = 12
Major_Feed_TargetWeight_Value(13) = 13
Major_Feed_TargetWeight_Value(14) = 14

FeedB_VolumeTarget_Value(1) = .2
FeedB_VolumeTarget_Value(2) = .2
FeedB_VolumeTarget_Value(3) = .3
FeedB_VolumeTarget_Value(4) = 1
FeedB_VolumeTarget_Value(5) = 3
FeedB_VolumeTarget_Value(6) = 0.0  
FeedB_VolumeTarget_Value(7) = 0.0
FeedB_VolumeTarget_Value(8) = 0.0
FeedB_VolumeTarget_Value(9) = 0.0
FeedB_VolumeTarget_Value(10) = 0.0
FeedB_VolumeTarget_Value(11) = 0.0
FeedB_VolumeTarget_Value(12) = 0.0
FeedB_VolumeTarget_Value(13) = 0.0
FeedB_VolumeTarget_Value(14) = 0.0


'Major Feed SetPoint (SP) variables
Dim Major_Feed_SP as Double = 250 '[mL/H] - starting pump flow setpoint 
Dim Major_Feed_Percent_Value as Double = 0.25 'reduce pump flow rate to 25% of SP 
Dim Major_Feed_SP_percent as Double = Major_Feed_SP * Major_Feed_Percent_Value  'stores the new reduced pump flow sp
Dim cut_pumpValue_Percent as Double = 0.75 'slow down when reached 75% of target value 
Dim Major_Feed_Counter as Double = 0 

'Glucose feed 
Dim Glucose_Feed_SP as Double = 150
Dim Glucose_Feed_Percent_Value as Double = 0.25 'reduce pump flow rate to 25% of SP 
Dim Glucose_Feed_SP_percent as Double = Glucose_Feed_SP * Glucose_Feed_Percent_Value  'stores the new reduced pump flow sp
Dim Glucose_Cut_pumpValue_Percent as Double = 0.75 'slow down when reached 75% of target value 

'Feed B 
Dim FeedB_FlowRate as Double = 15 '[mL/H]
Dim FeedB_Counter as Double = 0 '#
Dim FeedB_CounterIncrement as Double = 0
Dim FeedB_Totalizer as Double = 0'[mL]
Dim FeedB_Totalizer_Count as Double = 0 

'Define array - indexing starts at 0.
If s is Nothing then
Dim a(22) as Double
    a(1) = 0 'time zero bottle weight 
    a(2) = 0 'interval starting bottleweight 
    a(3) = 0 'current weight 
    a(4) = Major_Feed_TargetWeight
    a(5) = Major_Feed_Counter
    a(6) = Feed_StartTime 'initialized to 0
    a(7) = FeedB_Counter 'initialized to 0
    a(8) = FeedB_Totalizer 'initialized to 0
    a(9) = FeedB_VolumeTarget 'initialized to 0
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
    .IntJ = s(9) 'Track Feed B Fed, not true value
    .intK = s(2) 'static interval feed
    .intL = s(3) 'dynamic scale reading during phase 9
    '.intS = s(9)
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
        .logmessage("██████████████-----AUTOFEED Scivario Test Starting...")
        .phase = .phase + 1


'==========================================================================================
' PHASE 1 - Start Inoculation Timer
'==========================================================================================

    Case 1 'Start Inoculation Timer  


    If State <> lastState Then
        Select Case State
            Case 0
                .logalarm("██████ Scale is inactive ██████")
            Case 1
                .logalarm("██████ Scale is idle ██████")
            Case 2
                .logMessage("██████ Scale is busy ██████")
            Case 3
                .logalarm("██████ Scale is in manual mode ██████")
            Case 4
                .logalarm("██████ Scale is calibrating... ██████")
            Case 5
                .logalarm("██████ Scale has an error ██████")
        End Select
        lastState = State
    End If
    

    'set pump calibration values
    'If .FDCal <> pumpF_FCal Then
     '   .FDCal = pumpF_FCal
    'End If

    'If .FCCal <> pumpD_FCal Then
      '  .FCCal = pumpD_FCal
    'End If

    'If .FACal <> pumpE_FCal Then
     '   .FACal = pumpE_FCal
    'End If

    If .InoculationTime_H > 0 And (.runtime_H - .phaseStart_H) > 0.1 / 60 Then
        '.logmessage("██████████████----INOC TIMER STARTED FOR UNIT#: " & .unit & ". Pumps Calibrated ----")
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

            If .InoculationTime_H > Feed_Time(1) And (.runtime_H - .phaseStart_H) > 0.1 / 60 Then 'waits for first feed time + 6s delay;
                s(5) = s(5) + 1 'Major Feed Counter
                s(7) = s(7) + 1 'Feed B Counter, not currently used
                .phase = .phase + 1
            Elseif s(5) > MaxFeeds Then 'checks if all feeds are done
                .pumpEActive = False
                .pumpDActive = False
                .pumpFActive = False
                .logmessage("All Feeds Delivered: Major=" & CInt(s(5)) & "/" & MaxFeeds & ", Feed B=" & CInt(s(7)) & "/" & MaxFeeds & ". Stopping (Phase 16).")
                .phase = 16
            End If

        'reset logic, not currently used    
        Else If .OfflineM = 1 and .OfflineN > 0 and .OfflineO = 0 Then
            s(5) = .OfflineN + s(5) + 1
            s(7) = .OfflineN + s(7) + 1
            .phase = .phase + 1
        ElseIf .OfflineM = 1 and .OfflineN > 0 and .OfflineO = 1 Then 
            s(5) = s(5) + 1 'Major Feed Counter
            s(7) = s(7) + 1 'Feed B Counter    
            .phase = .phase + 1
        End If 

'==========================================================================================
' PHASE 3 - Assign New Values for Interval, looks for sample confirmation
'==========================================================================================

    Case 3 'assign new values for interval time and specified target weight to feed
        
        'major feed - pump C
        If s(5) = 1 Then 
            Feed_StartTime = Feed_Time(1) 's(6)
            Major_Feed_TargetWeight = Major_Feed_TargetWeight_Value(1)
       
        ElseIf s(5) = 2 Then 
            Feed_StartTime = Feed_Time(2) 
            Major_Feed_TargetWeight = Major_Feed_TargetWeight_Value(2)
            
        ElseIf s(5) = 3 Then ' 
            Feed_StartTime = Feed_Time(3)
            Major_Feed_TargetWeight = Major_Feed_TargetWeight_Value(3)
        
        ElseIf s(5) = 4 Then
            Feed_StartTime = Feed_Time(4)
            Major_Feed_TargetWeight = Major_Feed_TargetWeight_Value(4)

        ElseIf s(5) = 5 Then
            Feed_StartTime = Feed_Time(5)
            Major_Feed_TargetWeight = Major_Feed_TargetWeight_Value(5)
      
        ElseIf s(5) = 6 Then
            Feed_StartTime = Feed_Time(6)
            Major_Feed_TargetWeight = Major_Feed_TargetWeight_Value(6)
       
        ElseIf s(5) = 7 Then 
            Feed_StartTime = Feed_Time(7) 
            Major_Feed_TargetWeight = Major_Feed_TargetWeight_Value(7)

        ElseIf s(5) = 8 Then ' 
            Feed_StartTime = Feed_Time(8)
            Major_Feed_TargetWeight = Major_Feed_TargetWeight_Value(8)

        ElseIf s(5) = 9 Then
            Feed_StartTime = Feed_Time(9)
            Major_Feed_TargetWeight = Major_Feed_TargetWeight_Value(9)
 
        ElseIf s(5) = 10 Then
            Feed_StartTime = Feed_Time(10)
            Major_Feed_TargetWeight = Major_Feed_TargetWeight_Value(10)
   
        ElseIf s(5) = 11 Then
            Feed_StartTime = Feed_Time(11)
            Major_Feed_TargetWeight = Major_Feed_TargetWeight_Value(11)

        ElseIf s(5) = 12 Then
            Feed_StartTime = Feed_Time(12)
            Major_Feed_TargetWeight = Major_Feed_TargetWeight_Value(12) 

        ElseIf s(5) = 13 Then
            Feed_StartTime = Feed_Time(13)
            Major_Feed_TargetWeight = Major_Feed_TargetWeight_Value(13)

        ElseIf s(5) = 14 Then
            Feed_StartTime = Feed_Time(14)
            Major_Feed_TargetWeight = Major_Feed_TargetWeight_Value(14)
        End If 

        'Feed B | s(7) = feed B counter, s(8) = pump D totalizer, s(9) = Feed B volume target
        If s(7) = 1 Then 
            FeedB_Totalizer = .VFPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(1)
        ElseIf s(7) = 2 Then
            FeedB_Totalizer = .VFPV    
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(2)
        ElseIf s(7) = 3 Then
            FeedB_Totalizer = .VFPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(3)
        ElseIf s(7) = 4 Then
            FeedB_Totalizer = .VFPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(4)
        ElseIf s(7) = 5 Then
            FeedB_Totalizer = .VFPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(5)
        ElseIf s(7) = 6 Then
            FeedB_Totalizer = .VFPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(6)
        ElseIf s(7) = 7 Then
            FeedB_Totalizer = .VFPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(7)
        ElseIf s(7) = 8 Then
            FeedB_Totalizer = .VFPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(8)
        ElseIf s(7) = 9 Then
            FeedB_Totalizer = .VFPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(9)
        ElseIf s(7) = 10 Then
            FeedB_Totalizer = .VFPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(10)
        ElseIf s(7) = 11 Then
            FeedB_Totalizer = .VFPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(11)
        ElseIf s(7) = 12 Then
            FeedB_Totalizer = .VFPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(12)
        ElseIf s(7) = 13 Then
            FeedB_Totalizer = .VFPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(13)
        ElseIf s(7) = 14 Then
            FeedB_Totalizer = .VFPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(14)
        End If

        'Store feed parameters
        s(4) = Major_Feed_TargetWeight
        s(6) = Feed_StartTime
        s(8) = FeedB_Totalizer
        s(9) = FeedB_VolumeTarget
        
        '----MODIFIED: Removing O condition----
        'skips sample confirmation on day 0
        'If procDay <= 0 Then ' First decision tree, skips sample confirmation on day 0
        '    .intT = s(6) - .inoculationTime_H 'time to feed in hours.
        '    .phase = .phase + 2

        'if sample not confirmed before feed time, will sit and wait in next phase. this logmessage runs a day early
        If .intB <> procDay And .inoculationTime_H < (s(6) + sampleDelay) And ((.runtime_H - .PhaseStart_H) > 0.1 / 60) Then 'Waiting for intB to confirm sample
            .logwarning("██ - Phase: 3 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ". - ████ - Enter Glucose Target in internal A, then confirm sample in internal B fields.")
            .intT = s(6) - .inoculationTime_H 'time to feed in hours. 
            .phase = .phase + 1
        ElseIf .intB <> procDay And .inoculationTime_H > (s(6) + sampleDelay) Then 'if sample is not confirmed before phase 4, will skip glucose entry
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

    Case 4 'Waiting for sample confirmation

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
' PHASE 5 - Waiting for Feed B to Start
'==========================================================================================

    Case 5 'Waiting for feed B to start

        'Skips to phase 7 if no feed B is requested
        If ((.runtime_H - .PhaseStart_H) > 0.1/60) and .inoculationTime_H > s(6) and s(9) <> 0 Then '6 second delay to ensure values get stored to array. also accounts for hiccups in DWC, checks for fb entry
            .logmessage("Phase: 5.  Day: " & procDay & ".  Unit #: " & .unit & ".  ---Feed B Starting: " & s(9) & " [mL] to be fed.")
            .phase = .phase + 1
        ElseIf ((.runtime_H - .PhaseStart_H) > 0.1/60) and .inoculationTime_H > s(6) and s(9) <= 0 Then 
            .logmessage("Phase: 5.  Day: " & procDay & ".  Unit #: " & .unit & ".  ---No Feed B To be fed, skipping.")
            .phase = .phase + 2 'goes to phase 7
        End If

'==========================================================================================
' PHASE 6 - Feed B Start
'==========================================================================================
        
    Case 6 'turn on pump

        If .inoculationTime_H > s(6) Then 
            .pumpFActive = True 'turn feed B on 
        End If 
        's(7) = FeedB Counter, redundand with s(5)
        's(8) = FeedB_Totalizer: starting totalizer snapshot from phase 3
        's(9) = FeedB_VolumeTarget: volume chosen in phase 3

        If s(7) <= MaxFeeds Then 'checks if above max feeds -probably not necessary
            'if totalizer is less than target set in phase 3, keep feeding at specified flowrate
            If .VFPV < s(8) + s(9) Then
                .FFSP = FeedB_FlowRate
            'when totalizer exceeds target, turn feed B off and move to next phase
            Else 
            .logmessage("██ - Phase: 6 - Day " & procDay & " - Unit #: " & .unit & ". - ██ - Feed B Finished: " & s(9) & " [mL] Fed")
                .phase = .phase + 1
            End If 
        ElseIf s(7) > MaxFeeds AndAlso ((.runtime_H - .PhaseStart_H) > 0.1/60) Then '6 second delay to ensure values get stored to array. also accounts for hiccups in DWC
            .phase = 16 'finish him 
        End If 

'==========================================================================================
' PHASE 7 - Feed B Off, Next Step
'==========================================================================================

    Case 7  

        .pumpFActive = False 'feed B off
        
        'takes snapshot of scale weight after feed B, for feed A
        s(1) = Scale_Weight(1) 

        'max feed check - not necessary
        If .inoculationTime_H > Feed_Time(14) Then 
            .phase = 2 'finish him 
        Else
            .phase = .phase + 1 'continue loop 
        End If 
 
'==========================================================================================
' PHASE 8 - Prepare for feed A
'==========================================================================================

    Case 8 'Prepare for feed A 
    
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
        If .InoculationTime_H > s(6) AndAlso s(4) <> 0 AndAlso (.runtime_H - .phaseStart_H) > 0.1/60 Then 
            .phase = .phase + 1
        'If no feed A, skips to phase 10
	    ElseIf s(4) = 0 Then 
            .logmessage("Phase: 8.  Day: " & procDay & ".  Unit #: " & .unit & ".  ---No Feed A entered, skipping Feed A phase.")
		    .phase = .phase + 2
        End If

'==========================================================================================
' PHASE 9 - Feed A Runs
'==========================================================================================

    Case 9 'Feed A runs

        '*************PHASE VARIABLES*************

        'Dynamic feed A totalizer calculation, updates live during phase 9
        Dim FeedA_Fed_Totalizer As Double '[g]
        FeedA_Fed_Totalizer = .VDPV - s(13) '[g]

        'Dynamic feed A scale weight, updates live during phase 9
        s(3) = Scale_Weight(3) '[g] 

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
        If Abs(MajorFeed_Fed) > s(4) AndAlso (.runtime_H - .phaseStart_H) > 0.1/60 Then '(s(1) - s(2)) > s(12)
            .pumpDActive = False
            .intF = s(2) - s(3) 'Feed A Fed by scale, probably not necessary since updated live
            .logmessage("██ - Phase: 9 - Day " & procDay & " - Unit #: " & .unit & ". - ██ Feed A Fed:  " & formatNumber(MajorFeed_Fed, 2) & " [g]")
            .phase = .phase + 1
        'If dynamic totalizer is 110% above target weight, stop feed
        ElseIf (.runtime_H - .phaseStart_H) > 0.1/60 AndAlso FeedA_Fed_Totalizer > (1.1 * s(4)) 'if dynamic totalizer is above 110% of target weight then stop
            .pumpDActive = False
            .intF = FeedA_Fed_Totalizer 'record dynamic totalizer value
            .logalarm("██ - Phase: 9 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ".██ Feed A Totalizer: " & formatNumber(FeedA_Fed_Totalizer, 2) & " [g] is above 110% of target weight: " & s(4) & " [g]. Stopping Feed A.")
            .phase = .phase + 1
        End If

        'if need to reset loop, then set offline C = 2. when in phase 2, set offline C = 0 to assign phase, then once in phase 4 assign offline C = 1   
        If .OfflineO = 2 Then '0 = no error, 1 = resume normal count  , 2 = loop reset 
            .phase = 2
        End If

'==========================================================================================
' PHASE 10 - Glucose Decision Tree
'==========================================================================================

    Case 10 'Glucose Decision Tree

        s(10) = Scale_Weight(4) 'sets starting glucose interval scale weight
        s(11) = s(10) 'Initialize end weight to prevent negative display if feed is skipped
        s(12) = .intA 'requested glucose

        'Removing procday check
        'If procDay = 0 Then 'should never happen
        '    .logmessage("Phase: 10.  Day: " & procDay & ".  Unit #: " & .unit & ".  ---No Glucose addition needed, proceeding to Phase 13.")
        '    .phase = .phase + 5

        'intB = 0 means sample not confirmed
        'intB = 1 means sample confirmed
        '(s(6) + sampleDelay) = feed time + delay time

        If .intB <> procDay AndAlso .inoculationTime_H < (s(6) + sampleDelay) Then 'skips glucose if sample not confirmed
            .phase = .phase + 1
        ElseIf .intB = procDay Then
            .phase = .phase + 2
        End If

'==========================================================================================
' PHASE 11 - Wait for sample Confirmation
'==========================================================================================

    Case 11 'potential skip of glucose

        'intB = 0 means sample not confirmed
        'intB = 1 means sample confirmed
        '(s(6) + sampleDelay) = feed time + delay time
        s(12) = .intA 'requested glucose

        If .intB <> procDay 0 and .inoculationTime_H > (s(6) + sampleDelay) Then 'skips glucose if sample not confirmed
            .logwarning("██ - Phase: 11 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ". - ████ - Skipping glucose feed, sample not verified")
            s(11) = s(10) 'show glucose fed as 0
            .phase = .phase + 4
        ElseIf .intB = procDay Then 'if not entered until after feed
            .phase = .phase + 1
        End If

'==========================================================================================
' PHASE 12 - Record Glucose Phase Bottleweight Baseline and Totalizer Baseline
'==========================================================================================

    Case 12 'Record Glucose phase bottleweight and totalizer baseline, redundant with next phase

        s(10) = Scale_Weight(4) 'Static glucose interval starting scale weight
        s(11) = s(10) 'Initialize end weight to prevent negative display
        s(12) = .intA 'update requested glucose
        s(14) = .VEPV 'Record Glucose totalizer baseline at start of interval 
        
        If .InoculationTime_H > s(6) AndAlso (.runtime_H - .phaseStart_H) > 0.5/60 Then 'waits for respective feed time + 30s delay; 
            .phase = .phase + 1
        End If 

'==========================================================================================
' PHASE 13 - Prepare for Glucose
'==========================================================================================

    Case 13 'Prepare for Glucose Feed

        '*************PHASE VARIABLES*************

        'Record Glucose totalizer baseline at start of interval, should not update in phase 14
        s(14) = .VEPV '[mL]

        'Record Scale weight and set s(10) to starting weight for glucose
        s(10) = Scale_Weight(4) '[g]
        'Set dynamic weight to s(10) just in-case it were negative, will begin updating in phase 14
        s(11) = s(10) '[g]

        's(12) = user input glucose from phase 12

        '*************PHASE VARIABLES END*************

        If s(12) = 0 AndAlso (.runtime_H - .phaseStart_H) > 0.5/60 Then 'no glucose to be added, skips glucose feed
            .logmessage("Phase: 14.  Day: " & procDay & ".  Unit #: " & .unit & ".  ---Sampling Confirmed: " & s(12) & " [g] Glucose to be fed.")
            .phase = .phase + 2
        ElseIf s(12) <> 0 AndAlso (.runtime_H - .phaseStart_H) > 0.5/60 Then 
            .logmessage("Phase: 14.  Day: " & procDay & ".  Unit #: " & .unit & ".  ---Sampling Confirmed: " & s(12) & " [g] Glucose to be fed.")
            .phase = .phase + 1
        End If

        'If .VEPV <> 0 Then 'reset feed a totalizer
        '    .SetVEPV(0)
        'End If        

'==========================================================================================
' PHASE 14 - Glucose Feed
'==========================================================================================

    Case 14 'glucose feed 

        '*************PHASE VARIABLES*************

        'Dynamic Glucose totalizer calculation, updates live during phase 14
        Dim Glucose_Fed_Totalizer As Double '[g]
        Glucose_Fed_Totalizer = .VEPV - s(14)  '[g]

        'Dynamic glucose scale weight, updates live during phase 14
        s(11) = Scale_Weight(5) '[g]

        'intH is display for glucose fed, updates live during phase 14, phase 13 snapshot weight minus current weight
        .intH = s(10) - s(11) '[g] same as
        
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

        .pumpEActive = True 'Glucose starts

        'Speed Control Logic
        'if glucosefeedfed exceeds 75% of target weight, reduce flowrate to 25% of starting setpoint
        If Abs(glucoseFeed_Fed) > (s(12) * Glucose_Cut_pumpValue_Percent) Then 
            .FESP = Glucose_Feed_SP_percent 'reduced pump flow setpoint when 75% of target weight achieved 
        Else
            .FESP = Glucose_Feed_SP 'starting pump flow rate
        End If 
        
        '--- MODIFIED: Use totalizer-based fed amount ---
        If abs(glucoseFeed_Fed) > s(12) AndAlso (.runtime_H - .phaseStart_H) > 0.1/60 OrElse s(12) = 0 Then 'allows user to stop glucose value 
            .logmessage("Phase: 14  Day: " & procDay & ".  Unit #: " & .unit & ".  ---Feed Counter: [" & .intD &"]")
            .phase = .phase + 1 'reset to phase 2; need to move to glucose feed phase 

        ElseIf (.runtime_H - .phaseStart_H) > 0.1/60 AndAlso Glucose_Fed_Totalizer > (1.1 * s(12)) Then 'if dynamic totalizer is above 110% of target weight then stop
            .pumpEActive = False
            .logwarning("██ - Phase: 14 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ". - ████ - Totalizer: " & formatNumber(.VEPV, 2) & " [g] is above 110% of target weight: " & s(12) & " [g]. Stopping Feed A.")
            .phase = .phase + 1
        End If

        'if need to reset loop, then set offline C = 2. when in phase 2, set offline C = 0 to assign phase, then once in phase 4 assign offline C = 1   
        'If .offlineC = 2 Then '0 = no error, 1 = resume normal count  , 2 = loop reset 
        '    .logMessage("Need to reset counter. Set to phase 2")
        '    .phase = .phase + 1
        'End If

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

'Checks with logs - disable when not debugging
'p.LogMessage("Phase: " & p.phase)
'p.logmessage("FeedB Increment" & s(18))
'p.logmessage("static totalizer " & s(19))
'p.logmessage("feed b vol target " & s(20))
'p.logmessage("active totalizer " & p.VFPV)
'p.LogMessage("Inoc Time [h]: " & p.InoculationTime_H)
'p.LogMessage("Phase Interval Time [h]: " & (P.runtime_H-p.PhaseStart_H))

'p.LogMessage("Major Feed SP: " & Major_Feed_SP)
'p.LogMessage("pump C SP: " & p.FDSP)

'p.LogMessage("MajorFeed Fed: " & MajorFeed_Fed )
'p.logmessage("majorfeed target weight: " & Major_Feed_TargetWeight)
'p.LogMessage("MajorFeed Start Time: " & Major_Feed_StartTime)
'p.LogMessage("Array0 - time zero bw: " & s(0))
'p.LogMessage("Array1 - STATIC: " & s(1))
'p.LogMessage("Array2 - STATIC INTERVAL: " & s(2))
'p.LogMessage("Array3: -DYNAMIC INTERVAL" & s(3))
'p.LogMessage("Array4: " & s(4))
'p.LogMessage("Array5: " & s(5))
'p.LogMessage("Array6: " & s(6))
'p.LogMessage("Array7: " & s(7))
'p.LogMessage("array11: " & s(11))
'p.LogMessage("array10: " & s(10))
'p.LogMessage("array12: " & s(12))