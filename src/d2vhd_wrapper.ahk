; Ryan Paul ryunpaul[gmail] 02/02/17
; Ghetto Disk2vhd.exe custom partition selector
;
;
; Usage: d2vhd.ahk [/xstd?] [*|TERM...] OUTPUT_FILE
;  /x  Do not use Vhdx
;  /s  Do not use Volume Shadow Copy 
;  /t  Test mode (skip backup, no output file required)
;  /p  Keep backup results open after completion
;  /d  Debug mode (show debug info panel)
;  /?  Usage text

;   * will match all volumes ..OR..
;   TERM(s) will be matched against VOLUME and LABEL columns
;   (No TERM input defaults to both c:\ and "System Reserved")

; Ex: d2vhd.ahk /dt
; Ex: d2vhd.ahk * d:\allVolumeBackup.vhdx
; Ex: d2vhd.ahk e:\some\place\without\spaces\default_backup.vhdx
; Ex: d2vhd.ahk /x /s z:\ adultlabel "f:\my backup\with spaces\adultVolumes.vhd"
;
;
; Eight critical ste[s] for prepping and backing up
; 1) Command line parsing
; 2) Gather volume info
; 3) Reverse order of system volumes
; 4) Check volumes based on given user data before working with Disk2vhd
; 5) Early Exit Conditions
; 6) Start Disk2vhd.exe Gui
; 7) Enable desired volumes for backup
; 8) Manipulate backup options
; 9) Start the backup
; 10) Exit Behavior



if not A_IsAdmin {

    MsgBox,,, Needs Admin Privileges, 5
    ExitApp
}


disk2vhd_file := "Disk2vhd.exe"
if not (FileExist(disk2vhd_file)) {
    
    msgbox,,, % disk2vhd_file " not found!", 5
    ExitApp
}


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; 1) Command line parsing
; Get all them tastey switches from the command line
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; Look at command line arguments
argv := []
argc := getArgs(argv)

; Parse each argument
useVhdx := true
useShadowCopy := true
outputFile := ""
debugMode := false
testMode := false
allVolumeFlag := false
exitAfterBackup := true
defaultIncludeKeywords := ["C:\", "System Reserved"]
includeKeywords := []
for i, arg in argv {

    if (RegExMatch(arg, "^/")) {
        ; Switch flags

        chars := substr(arg, 2)

        for i, char in strsplit(chars) {
        
            ;vhdx
            if (char = "x")
                useVhdx := false

            ;shadow copy
            if (char = "s")
                useShadowCopy := false

            ;debug
            if (char = "d")
                debugMode := true

            ;test run
            if (char = "t")
                testMode := true

            ;test run
            if (char = "p")
                exitAfterBackup := false

            ;usage
            if (char = "?") {
                usage()
                ExitApp
            }
        }

    } else if (arg = "*") {
        ; Wildcard volumes
        
        allVolumeFlag := true

    } else if (RegExMatch(arg, "i)\.vhdx?$")) {
        ; Output file

        outputFile := arg

    } else {
        ; Term for including backup volume

        includeKeywords.push(arg)
    }
}

; If no arguments passed, fall back to defaults
includeKeywords := includeKeywords.Length() ? includeKeywords : defaultIncludeKeywords



;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; 2) Gather volume info
; See what treasures this system holds
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; Capture desired volumes
params := ["Automount", "BlockSize", "Capacity", "DeviceID", "DriveLetter", "DriveType", "FileSystem", "FreeSpace", "Label", "MaximumFileNameLength", "Name", "SerialNumber"]
query := "Select * from Win32_Volume"
volumeList := []
for volume in ComObjGet("winmgmts:").ExecQuery(query) {

    ; DRIVE_REMOVABLE (2), DRIVE_FIXED (3)
    ; https://msdn.microsoft.com/en-us/library/windows/desktop/aa364939(v=vs.85).aspx
    if ((volume.DriveType = 2) || (volume.DriveType = 3)) {

        obj := {}
        for i, param in params {
            obj[param] := volume[param]
        }

        volumeList.push(obj)
    }
}
volume := ""



;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; 3) Reverse order of system volumes
; Disk2vhd's list of system volumes is reveresed from what AHK queries. Must 
; reverse the system volumes to match disk2vhd's listview ordering.
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; Pull out system volumes (all identical Names starting with "\\?\...")
sysVols := []
idx := 1
while (idx <= volumeList.Length()) {

    vol := volumeList[idx]
    if RegExMatch(vol.Name, "\\\\?\\") {

        sysVols.push(vol)
        volumeList.RemoveAt(idx)
    } else {

        idx++
    }
}

names := []
for i,v in volumeList
    names.push(v.Name)

; Sort what's left (drive letters)
; Ghetto sort: object keys are auto sorted. (array -> object -> array)

; Part 1 (array -> object)
sortMe := {}
for i,vol in volumeList {
    sortMe[vol.Name] := vol
}

; Part 2 (object -> array)
volumeListSorted := []
for i,vol in sortMe {
    volumeListSorted.push(vol)
}

; Insert system volumes back into list (system vols will be reversed order)
for i,vol in sysVols {
    volumeListSorted.InsertAt(1, vol)
}



;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; 4) Check volumes based on given user data before working with Disk2vhd
; Decide if it's a viable solution
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; Volume properties to search through for filter keywords
; Name: "c:\" or "d:\", Label: "System Reserved" or "OS"
filterProperties := ["Name", "Label"]
backupFlag := false
loop, % volumeListSorted.Length()
{
    curVol := volumeListSorted[A_Index]
    searchState := true

    ; Add a property on our volume object
    curVol.D2VHD_BACKUP := 0

    ; Search through selected volume properties for filtering keyword
    while (searchState) {

        ; End of array; termination
        if (A_Index > filterProperties.Length())
            break

        prop := filterProperties[A_Index]

        for i,keyword in includeKeywords {

            if ((curVol[prop] = keyword) || (allVolumeFlag)) {

                curVol.D2VHD_BACKUP := true
                searchState := false
                backupFlag := true
                break
            }
        }        
    }
}

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; 5) Early Exit Conditions
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

if (debugMode) {

    DEBUG_ACTION()

    if not (testMode)
        return
}


if ((!outputFile) && (!testMode)) {

    MsgBox,,, No valid output file, 5
    exitApp
}


if not (backupFlag) {

    Msgbox,,, No matching volumes, 5
    ExitApp
}


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; 6) Start Disk2vhd.exe Gui
; Must be in the same directory as AHK script. Will close an existing
; Disk2vhd.exe window. Does not account for multiple prev windows.
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; Close if already open
WinClose, % "ahk_class Disk2VhdClass"

; Make sure EULA is set
regKey := "Software\Sysinternals\Disk2vhd"
RegWrite, REG_SZ, HKCU, % regKey, EulaAccepted, 1

; First remove any cached output file in registry
regKey := "Software\Sysinternals\Disk2Vhd"
RegWrite, REG_SZ, HKCU, % regKey, VhdFile, % ""

; Run Disk2vhd.exe and wait for window to appear
Run, % disk2vhd_file,,, winPID
WinWait, % "ahk_class Disk2VhdClass"



;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; 7) Enable desired volumes for backup
; Use tab selection to focus ListView control, then arrows to select
; the proper items in list. From there use space to unselect items.
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; Focusing the volume ListView is a bit wonky and done in two parts.
; 1) Tab jump one control past the ListView control
loop, 6
{
    Send, {Tab}
    Sleep, 100
}

; 2) Reverse tab jump to activate ListView control (?? profit)
Send, {Shift Down}{Tab}{Shift Up}

; loop, % volumeListSorted.Length() {
for idx, curVol in volumeListSorted {

    ; Select next item in d2vhd ListVIew
    Send, {Down}

    ; Disable if no include filtering keywords matched
    if not (curVol.D2VHD_BACKUP)
        Send, {Space}

    sleep, 50
}



;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; 8) Manipulate backup options
; Apply command line options
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; Output file
if (outputFile) {

    ; Focus from listview to output file input control
    Send {Shift Down}
    loop, 3
    {
        Send {Tab}
    }
    Send {Shift Up}
    Send % outputFile
}

; vhdx option
; toggling the vhdx checkbox will append or remove the 'x' on file path
send {Alt Down}
loop % useVhdx ? 2 : 3
{
    send s
    sleep 50
}
send {Alt Up}

; shadow copy option
if not (useShadowCopy)
    send {Alt Down}u{Alt Up}



;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; 9) Start the backup
; RIP
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; Last call
if not (testMode) {
    
    Send {Alt Down}c{Alt Up}

    sleep 100

    ; Destination not enough free space
    if WinExist("", "already exists")
        Send {Enter}

    ;if WinExist("", "not enough space")
    ;    Send {Enter}
}



;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; 10) Exit Behavior
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

if (exitAfterBackup) {

    SetTitleMatchMode, RegEx

    while(true) {

        WinGetText, completed, i)Disk2Vhd - Sysinternals, completed successfully

        if (completed) {

            WinClose, % "ahk_class Disk2VhdClass"
            break
        }
    }
}

ExitApp


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; Debug
; Collect misc data that was used to determine actions on Disk2vhd GUI
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

DEBUG_ACTION() {
    global

    ; Recreate disk2vhd listview interface
    desiredParams := ["Name", "Label", "Capacity", "FreeSpace"]
    gui, add, ListView, w400 r7, % join(desiredParams, "|")

    for i, curVol in volumeListSorted {

        colData := []
        for i, param in desiredParams {
            colData.push(curVol[param])
        }

        LV_Add("", colData*)
    }

    LV_ModifyCol()
    LV_ModifyCol(1, 100)


    ; Run through each volume collecting detailed info
    propertiesDump := []
    compareDump := []
    for idx, curVol in volumeListSorted {

        ; 1) Capture all properties of volume object
        propStrBuilder := ["[" curVol.Label "]"]
        for k,v in curVol
            propStrBuilder.push(k " = " v)

        propertiesDump.push(join(propStrBuilder, "`n") "`n")


        ; 2) Capture all the filtering tests
        filterStrBuilder := ["[Volume " idx "]"]
        for i,prop in filterProperties {
            
            ; Go over each keyword given in includeKeywords
            for i,keyword in includeKeywords {

                compString := curVol[prop] " = " keyword

                ; Does current volume property match current filter keyword?
                if (curVol[prop] = keyword)
                    compString := ">> " compString " <<"
            
                filterStrBuilder.push(compString)
            }
        }

        if (curVol.D2VHD_BACKUP)
            filterStrBuilder[1] := filterStrBuilder[1] "   ** Match **"

       
        separator := (A_Index < volumeListSorted.Length()) ? "`n" : ""
        compareDump.push(join(filterStrBuilder, "`n") separator)
    }


    ; Gather various app flags
    appFlags := ["useVhdx", "useShadowCopy", "debugMode", "testMode"]
    appFlagStrBld := []
    for i, v in appFlags
        appFlagStrBld.push(v " = " %v%)


    ; Show detailed info in edit box
    infoSections := ["-= FLAGS =-"
        , join(appFlagStrBld, "`n")
        , "`n"
        , "-= INCLUDE TERMS =-"
        , join(includeKeywords, "`n")
        , "`n"
        , "-= MATCH RESULTS =-"
        , join(compareDump, "`n")
        , "`n"
        , "-= VOLUME DATA =-"
        , join(propertiesDump, "`n")]

    winTitle := A_ScriptName " Debug Info"
    gui, add, Edit, Multi r30, % join(infoSections, "`n")
    gui, show,, % winTitle

    WinWaitClose, % winTitle
    if not (testMode)
        ExitApp
}



;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; Function Helpers
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

usage() {
    msg = 
(
Usage: %A_ScriptName% [/xstd?] [*|TERM...] OUTPUT_FILE
  /x  Do not use Vhdx
  /s  Do not use Volume Shadow Copy
  /t  Test mode (skip backup, no output file required)
  /p  Keep backup results open after completion
  /d  Debug mode (show debug info panel)
  /?  Usage text

  * will match all volumes ..OR..
  TERM(s) will be matched against VOLUME and LABEL columns
  (No TERM input defaults to both c:\ and "System Reserved")

Ex: d2vhd.ahk /dt
Ex: d2vhd.ahk * d:\allVolumeBackup.vhdx
Ex: d2vhd.ahk e:\some\place\without\spaces\default_backup.vhdx
Ex: d2vhd.ahk /x /s z:\ adultlabel "f:\my backup\with spaces\adultVolumes.vhd"
Ryan Paul
)

    DllCall("AttachConsole", "int", -1)
    FileAppend, % "`n" msg "`n", CONOUT$
}

join(a,d) {
    s := ""
    for i,v in a
        s .= v (i < a.length() ? d : "")
    return s
}

array_reverse(array, start:=0, end:="NULL") {

    len := array.Length()

    ; START: Adjust 1 based index, check signage, set defaults
    if (start > 0)
        idxFront := start - 1    ; Include starting index going forward
    else if (start < 0)
        idxFront := len + start  ; Count backwards from end
    else
        idxFront := start

    ; END: Check signage and set defaults
    if (end > 0)
        idxBack := end
    else if (end < 0)
        idxBack := len + end   ; Count backwards from end
    else
        idxBack := len

    idxFront := start
    idxBack := start + dist
    loop, % len - begin
    {
        tmp := array[idxFront]
        array[idxFront] := array[idxBack]
        array[idxBack] := tmp

        idxFront += 1
        idxBack -= 1
    }

    return array
}


; Add arguments to given array and return count
; Variables 0-n (yes, %0%) hold arg data. Variable 0 being count of args:
getArgs(ByRef argv) {
    global

    Loop, %0%
    {
        argv[A_Index] := %A_Index%
    }

    return %0%
}