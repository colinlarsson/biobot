'Abbvie | Dynamic time trigger and gravimetric response feed | JDL CL 03-13-2025
'Version 5.0 Last Update 2025-08-09

'A = Glucose gravimetric f.cal = 24
'B = Base not controlled
'C = Feed A gravimetric f.cal = 15.92
'D = Feed B volumetric f.cal = 105

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
Dim procDay As Integer 'Day used for logmessages and acknowledge glucose sample

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

'Feed A fcal (pump C) = 15.92
'Glucose fcal (pump A) = 24
'feed B fcal (pump D) = 105
Dim pumpC_FCal as Double
Dim pumpD_FCal as Double
Dim pumpA_FCal as Double

' How many scheduled events are configured (easy to change)
Dim MaxFeeds As Integer

pumpA_FCal = 24
pumpC_FCal = 15.92
pumpD_FCal = 105

MaxFeeds = 14

sampleDelay = 4

Feed_Time(1) = 24
Feed_Time(2) = 48
Feed_Time(3) = 72
Feed_Time(4) = 96
Feed_Time(5) = 120
Feed_Time(6) = 144
Feed_Time(7) = 200
Feed_Time(8) = 252
Feed_Time(9) = 500
Feed_Time(10) = 501
Feed_Time(11) = 502
Feed_Time(12) = 503
Feed_Time(13) = 504
Feed_Time(14) = 505

Major_Feed_TargetWeight_Value(1) = 20
Major_Feed_TargetWeight_Value(2) = 0
Major_Feed_TargetWeight_Value(3) = 15
Major_Feed_TargetWeight_Value(4) = 30
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

FeedB_VolumeTarget_Value(1) = 1.2
FeedB_VolumeTarget_Value(2) = 3
FeedB_VolumeTarget_Value(3) = 0
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
    s = a
End If

If P isNot Nothing Then
    With P 

    'interval starting weight minus current weight
    MajorFeed_Fed = (s(2) - s(3)) 
    'glucose feed starting weight minus current weight
    glucoseFeed_Fed = (s(10) - s(11)) 

    .intA = .phase 'phase
    .intB = .runtime_H - .phaseStart_H 'phase time 
    .intC = s(5) 'Feed Counter
    .intD = procDay 'feed b counter 
    .intE = s(4) 'major feed target weight  
    '.intF = s(2) - s(3) 'Feed A Fed Update:Defined later in feed A block
    .intG = s(12) 'glucose target
    .intH = s(10) - s(11) 'glucose fed
    .intI = s(9) + s(8) 'feed b volume target
    .IntJ = s(9) 'Track Feed B Fed, not true value
    .intK = s(2) 'static interval feed
    .intL = s(3) 'dynamic interval
    .intQ = procDay
    '.intS = s(9)
    '.intT = s(6) - .inoculationTime_H 'time to feed in hours. changing, won't be set until after phase 3 runs

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

Select Case .phase
    Case 0 'Workflow Start 
        .logmessage("██████████████-----AUTOFEED v.5.0 LAUNCHING...")
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
    
    'Feed A fcal (pump C) = 15.92
    'Glucose fcal (pump A) = 24
    'Feed B fcal (pump D) = 105

    'set pump calibration values
    If .FDCal <> pumpD_FCal Then
        .FDCal = pumpD_FCal
    End If

    If .FCCal <> pumpC_FCal Then
        .FCCal = pumpC_FCal
    End If

    If .FACal <> pumpA_FCal Then
        .FACal = pumpA_FCal
    End If

    If .InoculationTime_H > 0 And (.runtime_H - .phaseStart_H) > 0.1 / 60 Then
        '.logmessage("██████████████----INOC TIMER STARTED FOR UNIT#: " & .unit & ". Pumps Calibrated ----")
        .phase = .phase + 1
    End If
    
'==========================================================================================
' PHASE 2 - Counter Management and totalizer reset
'==========================================================================================

    Case 2
        .pumpCActive = False 
        .pumpDActive = False 
        .pumpAActive = False

        If .OfflineM = 0 Then 'Update Counters Manually, M: 1 enter reset state, N: number of steps to skip, O: step up 1. M will intercept after every feed
            If .InoculationTime_H > Feed_Time(1) And (.runtime_H - .phaseStart_H) > 0.1 / 60 Then 'waits for first feed time + 6s delay;
                s(5) = s(5) + 1 'Major Feed Counter
                s(7) = s(7) + 1 'Feed B Counter, not currently used
                .phase = .phase + 1
            Elseif s(5) > MaxFeeds Then 'checks if all feeds are done
                .pumpAActive = False
                .pumpCActive = False
                .pumpDActive = False
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
            FeedB_Totalizer = .VDPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(1)
        ElseIf s(7) = 2 Then
            FeedB_Totalizer = .VDPV    
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(2)
        ElseIf s(7) = 3 Then
            FeedB_Totalizer = .VDPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(3)
        ElseIf s(7) = 4 Then
            FeedB_Totalizer = .VDPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(4)
        ElseIf s(7) = 5 Then
            FeedB_Totalizer = .VDPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(5)
        ElseIf s(7) = 6 Then
            FeedB_Totalizer = .VDPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(6)
        ElseIf s(7) = 7 Then
            FeedB_Totalizer = .VDPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(7)
        ElseIf s(7) = 8 Then
            FeedB_Totalizer = .VDPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(8)
        ElseIf s(7) = 9 Then
            FeedB_Totalizer = .VDPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(9)
        ElseIf s(7) = 10 Then
            FeedB_Totalizer = .VDPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(10)
        ElseIf s(7) = 11 Then
            FeedB_Totalizer = .VDPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(11)
        ElseIf s(7) = 12 Then
            FeedB_Totalizer = .VDPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(12)
        ElseIf s(7) = 13 Then
            FeedB_Totalizer = .VDPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(13)
        ElseIf s(7) = 14 Then
            FeedB_Totalizer = .VDPV
            FeedB_VolumeTarget = FeedB_VolumeTarget_Value(14)
        End If

        'Store feed parameters
        s(4) = Major_Feed_TargetWeight
        s(6) = Feed_StartTime
        s(8) = FeedB_Totalizer
        s(9) = FeedB_VolumeTarget
        
        'skips sample confirmation on day 0
        If procDay <= 0 Then ' First decision tree, skips sample confirmation on day 0
            .intT = s(6) - .inoculationTime_H 'time to feed in hours.
            .phase = .phase + 2
        'if sample not confirmed before feed time, will sit and wait in next phase. this logmessage runs a day early
        ElseIf .offlineA <> procDay And .inoculationTime_H < (s(6) + sampleDelay) And ((.runtime_H - .PhaseStart_H) > 0.1 / 60) Then 'Waiting for offline A to confirm sample
            .logwarning("██ - Phase: 3 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ". - ████ - Enter Glucose Target in Offline Field, then set Day to the current Process Day.")
            .intT = s(6) - .inoculationTime_H 'time to feed in hours. 
            .phase = .phase + 1
        ElseIf .offlineA = procDay And .inoculationTime_H > s(6) 'if sample is confirmed before phase 4, this message will display, unlikely to happen
            .intT = s(6) - .inoculationTime_H 'time to feed in hours. 
            .logmessage("██ - Phase: 3 - Day " & procDay & " - Unit #: " & .unit & ". - ██ - Sampling Confirmed: Next Feeding Event in " & formatNumber(.intT, 2) & " Hrs")
            .logmessage("██ - Phase: 3 - Day " & procDay & " - Unit #: " & .unit & ". - ██ - Feed Targets: " & s(9) & " [mL] FB. " & S(4) & " [g] FA and " & .offlineB & " [g] Glucose.")
            .phase = .phase + 2
        End If

'==========================================================================================
' PHASE 4 - Waiting for Sample Confirmation
'==========================================================================================
'sits here after first feed, waiting for sample confirmation

    Case 4 'waiting for sample confirmation, goes to feed b once procDay is set

        If .offlineA <> procDay And .inoculationTime_H > (s(6) + sampleDelay) Then 'Waiting for offline A to confirm sample
            .logwarning("Phase 4: Sample not verified, feeding late, no glucose entered")
            .phase = .phase + 1
        ElseIf .offlineA = procDay and .inoculationTime_H > s(6) Then 'if sample is confirmed before phase 4, this message will display
            .logmessage("██ - Phase: 4 - Day " & procDay & " - Unit #: " & .unit & ". - ██ - Sampling Confirmed: Next Feeding Event in " & formatNumber(.intT, 2) & " Hrs")
            .logmessage("██ - Phase: 4 - Day " & procDay & " - Unit #: " & .unit & ". - ██ - Feed Targets: " & s(9) & " [mL] FB. " & S(4) & " [g] FA and " & .offlineB & " [g] Glucose.")
            .phase = .phase + 1
        End If

'==========================================================================================
' PHASE 5 - Waiting for Feed B to Start
'==========================================================================================

    Case 5 'Waiting for feed B to start

        If ((.runtime_H - .PhaseStart_H) > 0.1/60) and .inoculationTime_H > s(6) and s(9) <> 0 Then '6 second delay to ensure values get stored to array. also accounts for hiccups in DWC, checks for fb entry
            .logmessage("Phase: 5.  Day: " & procDay & ".  Unit #: " & .unit & ".  ---Feed B Starting: " & s(9) & " [mL] to be fed.")
            .phase = .phase + 1
        ElseIf ((.runtime_H - .PhaseStart_H) > 0.1/60) and .inoculationTime_H > s(6) and s(9) <= 0 Then 'corrected condition to check for no feed
            .logmessage("Phase: 5.  Day: " & procDay & ".  Unit #: " & .unit & ".  ---No Feed B To be fed, skipping.")
            .phase = .phase + 2 'goes to phase 7
        End If

'==========================================================================================
' PHASE 6 - Feed B Start
'==========================================================================================
        
    Case 6 'turn on pump

        If .inoculationTime_H > s(6) Then 
            .pumpDActive = True 'turn feed B on 
        End If 

        If s(7) <= MaxFeeds Then 'checks if above max feeds -probably not necessary
            If .VDPV <= s(8) + s(9) Then
                .FDSP = FeedB_FlowRate
            Else 'ends when totalize hits value
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

        .pumpDActive = False 'feed B off
        s(1) = Scale_Weight(1) 'static weight after feed B is finished 

        If .inoculationTime_H > Feed_Time(12) and s(4) > Major_Feed_TargetWeight_Value(12) Then 
            .phase = 2 'finish him 
        Else
            .phase = .phase + 1 'continue loop 
        End If 
 
'==========================================================================================
' PHASE 8 - Prepare for feed A
'==========================================================================================

    Case 8 'Prepare for feed A 

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

        If .VCPV <> 0 Then 'reset totalizer
            .SetVCPV(0)
        End If

        s(2) = Scale_Weight(2) 'Static of Interval scale weight 
        s(3) = s(2) 'Initialize end weight to prevent negative display if feed is skipped
        If .InoculationTime_H > s(6) AndAlso s(4) <> 0 AndAlso (.runtime_H - .phaseStart_H) > 0.1/60 Then 'waits for respective feed time + 6s delay; 
            .phase = .phase + 1
	    ElseIf s(4) = 0 Then 'if no feed A is entered, skip to glucose phase
            .logmessage("Phase: 8.  Day: " & procDay & ".  Unit #: " & .unit & ".  ---No Feed A entered, skipping Feed A phase.")
		.phase = .phase + 2
        End If

'==========================================================================================
' PHASE 9 - Feed A Runs
'==========================================================================================

    Case 9 'Feed A runs

        s(3) = Scale_Weight(3) 'dynamic interval scale weight 
        .pumpCActive = True 'replace with pump A for feed B 'issue here with offline reset 
        
        If abs(MajorFeed_Fed) > (s(4) * cut_pumpValue_Percent) Then  
            .FCSP = Major_Feed_SP_percent 'reduced pump flow setpoint 
        Else
            .FCSP = Major_Feed_SP 'starting pump flow rate
        End If 
        
        If abs(MajorFeed_Fed) > s(4) and (.runtime_H - .phaseStart_H) > 0.1/60 Then '(s(1) - s(2)) > s(12)
            .pumpCActive = False
            .intF = s(2) - s(3) 'Feed A Fed
            .logmessage("██ - Phase: 9 - Day " & procDay & " - Unit #: " & .unit & ". - ██ Feed A Fed:  " & formatNumber(MajorFeed_Fed, 2) & " [g]")
            .phase = .phase + 1
        ElseIf (.runtime_H - .phaseStart_H) > 0.1/60 AndAlso .VCPV > (1.1 * s(4)) 'if dynamic totalizer is above 110% of target weight then stop
            .pumpCActive = False
            .intF = .VCPV 'record dynamic totalizer value
            .logalarm("██ - Phase: 9 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ".██ Feed A Totalizer: " & formatNumber(.VCPV, 2) & " [g] is above 110% of target weight: " & s(4) & " [g]. Stopping Feed A.")
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

        If procDay = 0 Then 'should never happen
            .logmessage("Phase: 10.  Day: " & procDay & ".  Unit #: " & .unit & ".  ---No Glucose addition needed, proceeding to Phase 13.")
            .phase = .phase + 5
        ElseIf .offlineA <> procDay AndAlso .inoculationTime_H < (s(6) + sampleDelay) Then 'skips glucose if sample not confirmed
            '.logwarning("██ - Phase: 10 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ". - ████ - Enter Glucose Target in Offline Field, then set Day to the current Process Day.")
            .phase = .phase + 1
        ElseIf .offlineA = procDay Then
            .phase = .phase + 2
        End If

'==========================================================================================
' PHASE 11 - Check for sample Confirmation
'==========================================================================================

    Case 11 'potential skip of glucose

        If .offlineA <> procDay and .inoculationTime_H > (s(6) + sampleDelay) Then 'skips glucose if sample not confirmed
            .logwarning("██ - Phase: 11 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ". - ████ - Skipping glucose feed, sample not verified")
            .phase = .phase + 4
        ElseIf .offlineA = procDay Then 'if not entered until after feed
            .phase = .phase + 1
        End If

'==========================================================================================
' PHASE 12 - Record Glucose Phase Bottleweight
'==========================================================================================

    Case 12 'Record Glucose phase bottleweight and 

        s(10) = Scale_Weight(4) 'Static glucose interval starting scale weight
        s(11) = s(10) 'Initialize end weight to prevent negative display
        s(12) = .offlineB 'requested glucose 
        
        If .InoculationTime_H > s(6) and (.runtime_H - .phaseStart_H) > 0.1/60 Then 'waits for respective feed time + 6s delay; 
            .phase = .phase + 1
        End If 
'==========================================================================================
' PHASE 13 - Redundant Check
'==========================================================================================

    Case 13 'should skip over

        If s(12) = 0 AndAlso (.runtime_H - .phaseStart_H) > 0.1/60 Then
            '.logmessage("Phase: 13.  Day: " & procDay & ".  Unit #: " & .unit & ".  ---No glucose addition required on day: " & procDay & ". " & s(12) & " [g] was Fed")
            .phase = .phase + 2
        ElseIf s(12) <> 0 Then
            .phase = .phase + 1
        End If

        If .VAPV <> 0 Then 'reset feed a totalizer
            .SetVAPV(0)
        End If        

'==========================================================================================
' PHASE 14 - Glucose Feed
'==========================================================================================

    Case 14 'glucose feed 

        If .runtime_H - .phaseStart_H < 0.025/60 Then 
            .logmessage("Phase: 14.  Day: " & procDay & ".  Unit #: " & .unit & ".  ---Sampling Confirmed: " & s(12) & " [g] Glucose to be fed.")
        End If

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

        s(11) = Scale_Weight(5) 'dynamic interval scale weight 
        .pumpAActive = True 'Glucose starts
        
        If abs(glucoseFeed_Fed) > (s(12) * Glucose_Cut_pumpValue_Percent) Then  's(12) = glucose value input to offline D; cut value % = 75
            .FASP = Glucose_Feed_SP_percent 'reduced pump flow setpoint when 75% of target weight achieved 
        Else
            .FASP = Glucose_Feed_SP 'starting pump flow rate
        End If 
        
        If abs(glucoseFeed_Fed) > s(12) AndAlso (.runtime_H - .phaseStart_H) > 0.1/60 OrElse s(12) = 0 Then 'allows user to stop glucose value 
            .logmessage("Phase: 14  Day: " & procDay & ".  Unit #: " & .unit & ".  ---Feed Counter: [" & .IntC &"]")
            .phase = .phase + 1 'reset to phase 2; need to move to glucose feed phase 
        ElseIf (.runtime_H - .phaseStart_H) > 0.1/60 AndAlso .VAPV > (1.1 * s(12)) Then 'if dynamic totalizer is above 110% of target weight then stop
            .pumpAActive = False
            .logwarning("██ - Phase: 14 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ". - ████ - Totalizer: " & formatNumber(.VAPV, 2) & " [g] is above 110% of target weight: " & s(12) & " [g]. Stopping Feed A.")
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
        .pumpAActive = False
        If .runtime_H - .phaseStart_H < 0.1/60 Then 
        .logwarning("██ - Phase: 15 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ". - ████ - Total Feeds: - ██ - Feed A: " & formatNumber(MajorFeed_Fed, 2) & " [g]. Feed B: " & s(9) & " [mL]. Glucose:  " & formatNumber(glucoseFeed_Fed, 2) & "[g] - ██")
            .phase = 2
        End If

    Case 16 
    'insert stop commands 
    .pumpCActive = False
    .pumpDActive = False 
    .pumpAActive = False
End Select
End With
End If

'Checks with logs - disable when not debugging
'p.LogMessage("Phase: " & p.phase)
'p.logmessage("FeedB Increment" & s(18))
'p.logmessage("static totalizer " & s(19))
'p.logmessage("feed b vol target " & s(20))
'p.logmessage("active totalizer " & p.VDPV)
'p.LogMessage("Inoc Time [h]: " & p.InoculationTime_H)
'p.LogMessage("Phase Interval Time [h]: " & (P.runtime_H-p.PhaseStart_H))

'p.LogMessage("Major Feed SP: " & Major_Feed_SP)
'p.LogMessage("pump C SP: " & p.FCSP)

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