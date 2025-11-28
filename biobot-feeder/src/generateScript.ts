// src/generateScript.ts

export const generateVBScript = (config: any) => {
    const pumpProps: any = {
      A: { totalizer: 'VAPV', setpoint: 'FASP', active: 'pumpAActive', fcal: 'FACal', setPV: 'SetVAPV' },
      B: { totalizer: 'VBPV', setpoint: 'FBSP', active: 'pumpBActive', fcal: 'FBCal', setPV: 'SetVBPV' },
      C: { totalizer: 'VCPV', setpoint: 'FCSP', active: 'pumpCActive', fcal: 'FCCal', setPV: 'SetVCPV' },
      D: { totalizer: 'VDPV', setpoint: 'FDSP', active: 'pumpDActive', fcal: 'FDCal', setPV: 'SetVDPV' },
      E: { totalizer: 'VEPV', setpoint: 'FESP', active: 'pumpEActive', fcal: 'FECal', setPV: 'SetVEPV' },
      F: { totalizer: 'VFPV', setpoint: 'FFSP', active: 'pumpFActive', fcal: 'FFCal', setPV: 'SetVFPV' }
    };
  
    const feedB = pumpProps[config.feedBPump];
    const feedA = pumpProps[config.feedAPump];
    const glucose = pumpProps[config.glucosePump];
  
    // Helper to format arrays for VBA
    const formatArray = (arr: any[], prefix: string) => arr.map((v, i) => `${prefix}(${i + 1}) = ${v}`).join('\n');
  
    return `'Autofeed for Scivario - Generated Configuration
  'Version 5.5 - Generated: ${new Date().toISOString().split('T')[0]}
  'User: ${config.userName}
  'Experiment: ${config.experimentId}
  'Reactor: ${config.reactorId}
  
  '=== PUMP CONFIGURATION ===
  'Feed A (Major Feed): Pump ${config.feedAPump} - Gravimetric - FCal: ${config.feedAFCal}
  'Feed B: Pump ${config.feedBPump} - Volumetric - FCal: ${config.feedBFCal}
  'Glucose: Pump ${config.glucosePump} - Gravimetric - FCal: ${config.glucoseFCal}
  
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
  Dim calculatedProcDay As Integer
  calculatedProcDay = int((p.inoculationTime_H + 8) / 24)
  
  If calculatedProcDay < 0 Then
      procDay = 0
  Else
      procDay = calculatedProcDay
  End If
  
  'Configuration Parameters
  Dim sampleDelay as Double = ${config.sampleDelay}
  Dim Feed_StartTime as Double = 0
  Dim samplePromptLead_H As Double = 4
  
  'Feed Schedule Arrays
  Dim Feed_Time(15) as Double
  ${formatArray(config.schedule.map((s: any) => s.time), 'Feed_Time')}
  
  Dim Major_Feed_TargetWeight as Double
  Dim Major_Feed_TargetWeight_Value(15) as Double
  ${formatArray(config.schedule.map((s: any) => s.feedA), 'Major_Feed_TargetWeight_Value')}
  
  Dim FeedB_VolumeTarget as Double = 0
  Dim FeedB_VolumeTarget_Value(15) as Double
  ${formatArray(config.schedule.map((s: any) => s.feedB), 'FeedB_VolumeTarget_Value')}
  
  'Pump Calibrations
  Dim FeedA_FCal as Double = ${config.feedAFCal}
  Dim FeedB_FCal as Double = ${config.feedBFCal}
  Dim Glucose_FCal as Double = ${config.glucoseFCal}
  
  Dim MaxFeeds As Integer = ${config.schedule.length}
  
  'Feed A SetPoints
  Dim Major_Feed_SP as Double = ${config.feedAFlowRate}
  Dim Major_Feed_Percent_Value as Double = 0.25 
  Dim Major_Feed_SP_percent as Double = Major_Feed_SP * Major_Feed_Percent_Value
  Dim cut_pumpValue_Percent as Double = 0.75 
  Dim Major_Feed_Counter as Double = 0
  
  'Glucose Feed SetPoints
  Dim Glucose_Feed_SP as Double = ${config.glucoseFlowRate}
  Dim Glucose_Feed_Percent_Value as Double = 0.25
  Dim Glucose_Feed_SP_percent as Double = Glucose_Feed_SP * Glucose_Feed_Percent_Value
  Dim Glucose_Cut_pumpValue_Percent as Double = 0.75
  
  'Feed B Parameters
  Dim FeedB_FlowRate as Double = ${config.feedBFlowRate}
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
      MajorFeed_Fed = (s(2) - s(3)) 
      glucoseFeed_Fed = (s(10) - s(11)) 
  
      .intC = procDay 
      .intD = s(5) 
      .intE = s(4) 
      .intG = s(12) 
      .intH = s(10) - s(11) 
      .intI = s(9) + s(8) 
      .IntJ = .${feedB.totalizer} - s(8) 
      .intK = s(2) 
      .intL = s(3) 
      .intU = .phase 
      .intV = .runtime_H - .phaseStart_H 
  
      Dim timeRemaining As Double
      If s(5) = 0 Then 
          timeRemaining = Feed_Time(1) - .inoculationTime_H 
      Else 
          timeRemaining = s(6) - .inoculationTime_H 
      End If
  
      If timeRemaining < 0 Then .intT = 0 Else .intT = timeRemaining
      
      Select Case .phase
          Case 0 
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
  
                  .${feedA.active} = False
                  .${feedB.active} = False
                  .${glucose.active} = False
                  
                  If .OfflineM = 0 Then 
                      If .InoculationTime_H > Feed_Time(1) And (.runtime_H - .phaseStart_H) > 0.1/60 Then 
                          s(5) = s(5) + 1
                          s(7) = s(7) + 1
                          .phase = .phase + 1
                      ElseIf s(5) > MaxFeeds Then 
                          .${feedA.active} = False
                          .${feedB.active} = False
                          .${glucose.active} = False
                          .logmessage("All Feeds Delivered: Major=" & CInt(s(5)) & "/" & MaxFeeds & ", Feed B=" & CInt(s(7)) & "/" & MaxFeeds & ". Stopping (Phase 16).")
                          .phase = 16
                      End If
                  End If
   
  '==========================================================================================
  ' PHASE 3 - Assign New Values for Interval, looks for sample confirmation
  '==========================================================================================
  
              Case 3
                  If s(5) >= 1 And s(5) <= 14 Then
                      Feed_StartTime = Feed_Time(s(5))
                      Major_Feed_TargetWeight = Major_Feed_TargetWeight_Value(s(5))
                  End If
                  
                  If s(7) >= 1 And s(7) <= 14 Then
                      FeedB_Totalizer = .${feedB.totalizer}
                      FeedB_VolumeTarget = FeedB_VolumeTarget_Value(s(7))
                  End If
                  
                  s(4) = Major_Feed_TargetWeight
                  s(6) = Feed_StartTime
                  s(8) = FeedB_Totalizer
                  s(9) = FeedB_VolumeTarget
                  
                  If .intB <> procDay And .inoculationTime_H < (s(6) + sampleDelay) And ((.runtime_H - .PhaseStart_H) > 0.1 / 60) Then 
                      .logwarning("██ - Phase: 3 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ". - ████ - Enter Glucose Target in internal A, then confirm sample in internal B fields.")
                      .intT = s(6) - .inoculationTime_H 
                      .phase = .phase + 1
                  ElseIf .intB <> procDay And .inoculationTime_H > (s(6) + sampleDelay) Then 
                      .logwarning("██ - Phase: 3 - Day " & procDay & " - Unit #: " & .unit & ". - ██ - Sample not confirmed, feeding late, no glucose entered")
                      .intT = s(6) - .inoculationTime_H 
                      .phase = .phase + 2
                  ElseIf .intB = procDay And .inoculationTime_H > s(6) 
                      .intT = s(6) - .inoculationTime_H 
                      .logmessage("██ - Phase: 3 - Day " & procDay & " - Unit #: " & .unit & ". - ██ - Sampling Confirmed: Next Feeding Event in " & formatNumber(.intT, 2) & " Hrs")
                      .logmessage("██ - Phase: 3 - Day " & procDay & " - Unit #: " & .unit & ". - ██ - Feed Targets: " & s(9) & " [mL] FB. " & S(4) & " [g] FA and " & .intA & " [g] Glucose.")
                      .phase = .phase + 2
                  End If
  
  '==========================================================================================
  ' PHASE 4 - Waiting for Sample Confirmation
  '==========================================================================================
  
              Case 4
                  s(12) = .intA 
                  If .intB <> procDay And .inoculationTime_H > (s(6) + sampleDelay) Then 
                      .logwarning("Phase 4: Sample not verified, feeding late, no glucose entered")
                      .phase = .phase + 1
                  ElseIf .intB = procDay and .inoculationTime_H > s(6) Then 
                      .logmessage("██ - Phase: 4 - Day " & procDay & " - Unit #: " & .unit & ". - ██ - Sampling Confirmed: Next Feeding Event in " & formatNumber(.intT, 2) & " Hrs")
                      .logmessage("██ - Phase: 4 - Day " & procDay & " - Unit #: " & .unit & ". - ██ - Feed Targets: " & s(9) & " [mL] FB. " & S(4) & " [g] FA and " & .intA & " [g] Glucose.")
                      .phase = .phase + 1
                  End If
  
  '==========================================================================================
  ' PHASE 5 - Determines if Feed B is requested
  '==========================================================================================
  
              Case 5
                  If ((.runtime_H - .PhaseStart_H) > 0.1/60) And .inoculationTime_H > s(6) And s(9) <> 0 Then 
                      .logmessage("Phase: 5.  Day: " & procDay & ".  Unit #: " & .unit & ".  ---Feed B Starting: " & s(9) & " [mL] to be fed.")
                      .phase = .phase + 1
                  ElseIf ((.runtime_H - .PhaseStart_H) > 0.1/60) And .inoculationTime_H > s(6) And s(9) <= 0 Then 
                      .logmessage("Phase: 5.  Day: " & procDay & ".  Unit #: " & .unit & ".  ---No Feed B To be fed, skipping.")
                      .phase = .phase + 2 
                  End If
  
  '==========================================================================================
  ' PHASE 6 - Feed B Start
  '==========================================================================================                
  
              Case 6
                  If .inoculationTime_H > s(6) Then
                      .${feedB.active} = True 
                  End If
  
                  If s(7) <= MaxFeeds Then
                      If .${feedB.totalizer} < s(8) + s(9) Then
                          .${feedB.setpoint} = FeedB_FlowRate
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
                  .${feedB.active} = False
                  s(1) = Scale_Weight(1)
                  If .inoculationTime_H > Feed_Time(14) Then
                      .phase = 2
                  Else
                      .phase = .phase + 1
                  End If
   
  '==========================================================================================
  ' PHASE 8 - Prepare for feed A
  '==========================================================================================
  
              Case 8
                  s(13) = .${feedA.totalizer}
                  s(2) = Scale_Weight(2) 
                  s(3) = s(2) 
                  
                  If .inoculationTime_H > s(6) And s(4) <> 0 And (.runtime_H - .phaseStart_H) > 0.1/60 Then
                      .phase = .phase + 1
                  ElseIf s(4) = 0 Then
                      .logmessage("Phase: 8.  Day: " & procDay & ".  Unit #: " & .unit & ".  No Feed A entered, skipping Feed A phase.")
                      .phase = .phase + 2
                  End If
  
  '==========================================================================================
  ' PHASE 9 - Feed A Runs
  '==========================================================================================
  
              Case 9
                  Dim FeedA_Fed_Totalizer As Double
                  FeedA_Fed_Totalizer = .${feedA.totalizer} - s(13) 
                  s(3) = Scale_Weight(3)
                  .intF = s(2) - s(3) 
                  .${feedA.active} = True
                  
                  If Abs(MajorFeed_Fed) > (s(4) * cut_pumpValue_Percent) Then
                      .${feedA.setpoint} = Major_Feed_SP_percent 
                  Else
                      .${feedA.setpoint} = Major_Feed_SP 
                  End If
                  
                  If Abs(MajorFeed_Fed) > s(4) And (.runtime_H - .phaseStart_H) > 0.1/60 Then 
                      .${feedA.active} = False
                      .intF = s(2) - s(3) 
                      .logmessage("██ - Phase: 9 - Day " & procDay & " - Unit #: " & .unit & ". - ██ Feed A Fed:  " & formatNumber(MajorFeed_Fed, 2) & " [g]")
                      .phase = .phase + 1
                  ElseIf (.runtime_H - .phaseStart_H) > 0.1/60 And FeedA_Fed_Totalizer > (1.1 * s(4)) Then 
                      .${feedA.active} = False
                      .intF = FeedA_Fed_Totalizer 
                      .logalarm("██ - Phase: 9 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ".██ Feed A Totalizer: " & formatNumber(FeedA_Fed_Totalizer, 2) & " [g] is above 110% of target weight: " & s(4) & " [g]. Stopping Feed A.")
                      .phase = .phase + 1
                  End If
                  If .OfflineO = 2 Then .phase = 2
  
  '==========================================================================================
  ' PHASE 10 - Glucose Decision Tree
  '==========================================================================================
  
              Case 10
                  s(10) = Scale_Weight(4) 
                  s(11) = s(10) 
                  s(12) = .intA 
                    If .intB <> procDay AndAlso .inoculationTime_H > s(6) Then 
                        .phase = .phase + 1 
                    ElseIf .intB = procDay Then
                        .phase = .phase + 2
                    End If
  
  '==========================================================================================
  ' PHASE 11 - Wait for sample Confirmation
  '==========================================================================================
  
              Case 11
                  s(12) = .intA 
                  If .intB <> procDay and .inoculationTime_H > (s(6) + sampleDelay) Then 
                      .logwarning("██ - Phase: 11 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ". - ████ - Skipping glucose feed, sample not verified")
                      s(11) = s(10) 
                      .phase = .phase + 4
                  ElseIf .intB = procDay Then 
                      .phase = .phase + 1
                  End If
  
  '==========================================================================================
  ' PHASE 12 - Record Glucose Phase Bottleweight Baseline and Totalizer Baseline
  '==========================================================================================
  
              Case 12
                  s(10) = Scale_Weight(4) 
                  s(11) = s(10) 
                  s(12) = .intA 
                  s(14) = .${glucose.totalizer} 
                  If .inoculationTime_H > s(6) And (.runtime_H - .phaseStart_H) > 0.5/60 Then 
                      .phase = .phase + 1
                  End If
  
  '==========================================================================================
  ' PHASE 13 - Prepare for Glucose
  '==========================================================================================
  
              Case 13
                  s(14) = .${glucose.totalizer} 
                  s(10) = Scale_Weight(4) 
                  s(11) = s(10) 
                  If s(12) = 0 AndAlso (.runtime_H - .phaseStart_H) > 0.5/60 Then 
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
                  Dim Glucose_Fed_Totalizer As Double 
                  Glucose_Fed_Totalizer = .${glucose.totalizer} - s(14) 
                  s(11) = Scale_Weight(5) 
                  .intH = s(10) - s(11) 
                  .${glucose.active} = True
                  
                  If Abs(glucoseFeed_Fed) > (s(12) * Glucose_Cut_pumpValue_Percent) Then
                      .${glucose.setpoint} = Glucose_Feed_SP_percent
                  Else
                      .${glucose.setpoint} = Glucose_Feed_SP
                  End If
                  
                  If Abs(glucoseFeed_Fed) > s(12) And (.runtime_H - .phaseStart_H) > 0.1/60 Then
                      .logmessage("Phase: 14  Day: " & procDay & ".  Unit #: " & .unit & ".  ---Feed Counter: [" & .intD &"]")
                      .phase = .phase + 1
                  ElseIf (.runtime_H - .phaseStart_H) > 0.1/60 And Glucose_Fed_Totalizer > (1.1 * s(12)) Then
                      .${glucose.active} = False
                      .logwarning("██ - Phase: 14 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ". - ████ - Totalizer: " & formatNumber(.VEPV, 2) & " [g] is above 110% of target weight: " & s(12) & " [g]. Stopping Feed A.")
                      .phase = .phase + 1
                  End If
  
  '==========================================================================================
  ' PHASE 15 - Wrap Up
  '==========================================================================================
  
              Case 15
                  .${glucose.active} = False
                  If .runtime_H - .phaseStart_H > 0.1/60 Then
                      .logwarning("██ - Phase: 15 - ██ - Day: " & procDay & " - ██ - Unit #: " & .unit & ". - ████ - Total Feeds: - ██ - Feed A: " & formatNumber(MajorFeed_Fed, 2) & " [g]. Feed B: " & s(9) & " [mL]. Glucose:  " & formatNumber(glucoseFeed_Fed, 2) & "[g] - ██")
                      .phase = 2
                  End If
                  
              Case 16
                  .${feedA.active} = False
                  .${feedB.active} = False
                  .${glucose.active} = False
                  
          End Select
      End With
  End If
  `;
  };