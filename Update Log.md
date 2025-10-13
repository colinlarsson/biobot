## 2025-06-05 Update

Line: 500 
    - added final logwarning
Line: 26
    - Dim sampleDelay as Double = 8 'Set to 2/60 for testing 
    - Created to change delay wait time
Line: 335 - sample delay
Line: 315 - sample delay
Line: 440 - sample delay

Note: Considering andAlso Formatting


## 2025-06-06 Update

Line: 537-538
    - When glucose was skipped due to sample day, it skipped to phase 14, changed to 15.

## 2025-06-13 Update

Line: 210
    - Added fcal automatic updates, need to give option in form.

## 2025-06-20 Updates 4.17

Line: 523
    - fixed logmessage for feed A, read wrong phase

Line: 580 - fix for glucose calculation
    - takes reading of bottleweight s(10) to properly calculate glucose fed if feeding is skipped over. will get overwritten if not skipped.

Line: 41 - *Fix for process day counter*
    - now should read the correct process day at start of the day, doesn't wait for inoculation time

Line: 405 - *Fix for process day to read correctly when sample isn't verified.
    -   changed to procDay + 1

Line: 252 and 66 *added pump f.cal values for form functionality
    - pumpC_FCal = 15.92
    - pumpD_FCal = 24
    - pumpA_FCal = 105

Lines: 410 - 424 *fixed incorrect time to feed reading
    - inT (time to feed) not set until phase 3
    - .intT = s(6) - .inoculationTime_H 'time to feed in hours. 
    - need to verify it has time to update

## 2025-06-24 Updates 4.17

Lines: 40 - 50
    - Added calculated process day to keep procDay from going negative

Lines: 200
    - Moved .inT to display after phase 3
