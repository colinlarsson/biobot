' Name: Capacitance Based Temp Shift for Project 383 with intY
' Description: This script shifts the temperature setpoint based on capacitance criteria and time limits, using intY to track stages.
' Author: Brendan McGuire, Colin Larsson
' Updates: 2025-06-17 added rampstart and waitstart to track time in each stage CL
' Updated: 2025-06-18 array a(22) to leverage persistent array s(22) to keep rampStart and waitStart persistent CL
' Updated: 2025-06-23 removed waitstart and rampstart

dim capCriteria as double = 10.8  ' capacitance target
dim timeLimit as double = 68  ' hour limit on the shift
dim initialTemp as double = 36 'initial temperature
dim tempTarget as double = 32  'temp to shift to
dim rampDuration as double = 2   'ramp the temp shift over time in hours
dim capOverTime as double = 0.5 'in hours, how long must the capacitance be over the limit?
dim rampEndTime as double = 1 ' empty for now but will contain the end time of the ramping
dim tempDiff as double = initialTemp - tempTarget
dim rampPercent as double = 50 ' how far along the ramp are we so far
'dim waitStart as double = 0
'dim rampStart as double = 0
'dim .intY as integer ' this is the stage of the temp shift, 0 = initial, 1 = wait, 2 = ramping, 3 = done


If s is Nothing then
    Dim a(22) as Double
    a(1) = 0 'time zero bottle weight used for differenct script
    a(2) = 0 'interval starting bottleweight used for differenct script
    a(3) = 0 'current weight used for differenct script
    a(4) = Major_Feed_TargetWeight 'used for differenct script
    a(5) = Major_Feed_Counter 'used for differenct script
    a(6) = Feed_StartTime 'initialized to 0 used for differenct script
    a(7) = FeedB_Counter 'initialized to 0 used for differenct script
    a(8) = FeedB_Totalizer 'initialized to 0 used for differenct script
    a(9) = FeedB_VolumeTarget 'initialized to 0 used for differenct script
    a(10) = 0 'Glucose interval starting bottleweight used for differenct script
    a(11) = 0 'current weight used for differenct script
    a(12) = 0 'user input glucose used for differenct script
    a(15) = 0 ' Timestamp when capacitance wait period begins
    a(16) = 0 ' Timestamp when temperature ramp begins
    s = a ' s is a persistent variable that holds the state of the script
End If


If P IsNot Nothing Then
    With P

    select case .intY
	case 0 'initialize
		if .extA < capCriteria and .inoculationtime_h < timeLimit then
			.TSP = initialTemp
		elseif .extA > capCriteria then
			.intY = 1
            s(15) = .inoculationtime_h ' CRITICAL: Persist waitStart time
			.logmessage("Capacitance Criteria Met. Waiting " & (capOvertime * 60) & " minutes for confirmation.")
			.logmessage("Wait start time: " & s(15))
		elseif .inoculationtime_h >= timeLimit then
			.intY = 2
            s(16) = .inoculationtime_h ' CRITICAL: Persist rampStart time
            .logmessage("Time Limit Reached, Starting Temp Ramp")

		end if
	case 1 'wait time to see if capacitance is still good
		' s(15) > 0 check ensures we don't trigger on initialization
		if .extA > capCriteria and s(15) > 0 and .inoculationtime_h - s(15) >= capOvertime then
			.intY = 2
            s(16) = .inoculationtime_h ' CRITICAL: Persist rampStart time
            .logmessage("Capacitance Criteria stable, Starting Temp Ramp")
            .logmessage("Rampstart: " & s(16))
		elseif .extA < capCriteria then
			.intY = 0
            s(15) = 0 ' Reset wait timer since criteria is no longer met
            .logmessage("Capacitance Criteria Not Met, Returning to Initial State") 
		end if
	case 2 'the criteria has been met, start ramping
		' s(16) > 0 check ensures we don't trigger on initialization
		if s(16) > 0 and .inoculationtime_h - s(16) < rampDuration then
			rampPercent = ( .inoculationtime_h - s(16) ) / rampDuration
			.TSP = initialTemp - (rampPercent * tempDiff)
		elseif s(16) > 0 and .inoculationtime_h - s(16) >= rampDuration then	
			.TSP = tempTarget
			' Reset the ramp start time since ramping is complete
			.intY = 3
			.logmessage("Ramping complete, maintaining target temperature")
		end if
	' Final stage: Maintain the target temperature once ramping is complete
	case 3
			.TSP = tempTarget
			' No further action needed, just maintain the target temperature
			.logmessage("Currently maintaining target temperature: " & tempTarget & "°C")
	end select

    End With
End If
