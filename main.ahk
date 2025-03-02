#Requires AutoHotkey v2.0

; -----------------------------------------------------------------------------
; Retrieves image dimensions using GDI+.
; Returns true if width and height are both > 0, else false.
ImageGetSize(filePath, &width, &height) {
    global pToken
    Local pBitmap
    width := 0, height := 0
    if (!pToken) {
        LogAction("GDI+ not initialized!")
        return false
    }
    pBitmap := Gdip_CreateBitmapFromFile(filePath)
    if (pBitmap) {
        width := Gdip_GetImageWidth(pBitmap)
        height := Gdip_GetImageHeight(pBitmap)
        Gdip_DisposeImage(pBitmap)
    } else {
        LogAction("Failed to load image: " . filePath)
    }
    return (width > 0 && height > 0)
}

; -----------------------------------------------------------------------------
; Enhanced SaveScreenshot function with timestamp
SaveScreenshot(filePath) {
    global pToken
    Local X, Y, W, H, pBitmap
    Local timestamp := FormatTime(, "yyyyMMdd_HHmmss")
    
    ; Add timestamp to the filename if not already in the path
    if (!InStr(filePath, timestamp)) {
        SplitPath(filePath, &fileName, &fileDir, &fileExt)
        filePath := fileDir . "\" . SubStr(fileName, 1, InStr(fileName, ".", , , 1) - 1) . "_" . timestamp . "." . fileExt
    }
    
    if (!pToken) {
        LogAction("GDI+ not initialized!")
        return false
    }
    ; Get active window coordinates and capture screenshot
    WinGetPos(&X, &Y, &W, &H, "A")
    LogAction("Capturing screenshot of window at X=" . X . ", Y=" . Y . ", W=" . W . ", H=" . H)
    pBitmap := Gdip_BitmapFromScreen(X "|" Y "|" W "|" H)
    
    if (pBitmap) {
        Gdip_SaveBitmapToFile(pBitmap, filePath)
        Gdip_DisposeImage(pBitmap)
        LogAction("Screenshot saved to: " . filePath)
        return true
    } else {
        LogAction("Failed to create bitmap for screenshot")
        return false
    }
}

; -----------------------------------------------------------------------------
; Helper function to create backup of upgrade_button.png
BackupAndCaptureUpgradeButton() {
    global buttonImageFile, pToken
    Local sourceFile, backupFile, scriptDir
    
    scriptDir := A_ScriptDir
    sourceFile := scriptDir . "\" . buttonImageFile
    
    if (FileExist(sourceFile)) {
        ; Create backup with timestamp
        timestamp := FormatTime(, "yyyyMMdd_HHmmss")
        backupFile := scriptDir . "\backup_" . SubStr(buttonImageFile, 1, InStr(buttonImageFile, ".", , , 1) - 1) . "_" . timestamp . ".png"
        FileCopy(sourceFile, backupFile, 1)
        LogAction("Created backup of " . buttonImageFile . " at " . backupFile)
    }
    
    ; Prepare to capture new upgrade button image
    MsgBox("Position your mouse over the Upgrade button and press OK. The current screen area around your cursor will be captured as the new upgrade button reference image.", "Capture Upgrade Button")
    
    ; Get mouse position
    MouseGetPos(&mouseX, &mouseY)
    
    ; Capture a region around the mouse cursor
    captureWidth := 200
    captureHeight := 60
    captureX := mouseX - captureWidth / 2
    captureY := mouseY - captureHeight / 2
    
    ; Using GDI+ to capture the region
    pBitmap := Gdip_BitmapFromScreen(captureX . "|" . captureY . "|" . captureWidth . "|" . captureHeight)
    
    if (pBitmap) {
        Gdip_SaveBitmapToFile(pBitmap, sourceFile)
        Gdip_DisposeImage(pBitmap)
        LogAction("New upgrade button image captured and saved as " . sourceFile)
        MsgBox("Successfully captured new upgrade button image!", "Success")
        return true
    } else {
        LogAction("Failed to capture new upgrade button image")
        MsgBox("Failed to capture the upgrade button image. Please try again.", "Error")
        return false
    }
}

; -----------------------------------------------------------------------------
; Displays a message box with a countdown timer.
TimedMsgBox(title, message, timeout) {
    global timedMsgBoxActive, timedMsgBoxEndTime
    timedMsgBoxActive := true
    timedMsgBoxEndTime := A_TickCount + (timeout * 1000)
    SetTimer(UpdateTimedMsgBox, 500)
    MsgBox(message, title)
    SetTimer(UpdateTimedMsgBox, 0)
    timedMsgBoxActive := false
}

; Updates the title of the active window with the remaining countdown.
UpdateTimedMsgBox() {
    global timedMsgBoxActive, timedMsgBoxEndTime
    Local remaining
    if (!timedMsgBoxActive)
        return
    remaining := (timedMsgBoxEndTime - A_TickCount) / 1000
    if (remaining <= 0) {
        SetTimer(UpdateTimedMsgBox, 0)
        WinClose("A")
        return
    }
    WinSetTitle("Waiting... " . Round(remaining) . "s remaining", "A")
}

; -----------------------------------------------------------------------------
; Logs a timestamped message to file and debug output.
LogAction(message) {
    Local timestamp
    timestamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    try {
        FileAppend(timestamp . " - " . message . "`n", "IdleonAutomation.log")
    } catch {
        ToolTip("Error logging to file", 100, 200)
        Sleep(2000)
        ToolTip()
    }
    OutputDebug(timestamp . " - " . message)
}

; -----------------------------------------------------------------------------
; On script exit: stops batch processing (if active) and logs cleanup.
; ExitFunc(ExitReason, ExitCode) {
;     global pToken, batchProcessingActive
;     if (batchProcessingActive)
;         StopBatchProcessing()
;     ; Note: We no longer shut down pToken here because it's reused
;     LogAction("Script terminated and resources cleaned up")
; }

; -----------------------------------------------------------------------------
; Loads UI element positions and batch settings from the configuration file.
LoadPositionsFromFile() {
    global capturedPositions, configFile, lastIconSequence, lastRepetitionCount, lastWaitTimeBetweenCycles
    Local fileSize, fileContent, lines, lineCount, positionsFound, i, line, pos, key, value, coords
    Local inPositionsSection := false, inBatchSettingsSection := false
    
    LogAction("Attempting to load positions from: " . configFile)
    if (!FileExist(configFile)) {
        LogAction("Config file does not exist: " . configFile)
        try {
            FileAppend("[Positions]`n[BatchSettings]`n", configFile)
            if (FileExist(configFile)) {
                LogAction("Created new empty config file: " . configFile)
                MsgBox("Created new configuration file: " . configFile)
            } else {
                LogAction("ERROR: Failed to create config file")
                MsgBox("Warning: Cannot create files in the script directory. Check permissions.")
            }
        } catch {
            LogAction("ERROR: Exception when creating config file")
            MsgBox("Error creating configuration file. Check permissions.")
        }
        capturedPositions := Map()
        lastIconSequence := ""
        lastRepetitionCount := 1
        lastWaitTimeBetweenCycles := 5
        return
    }
    
    try {
        fileSize := FileGetSize(configFile)
    } catch {
        fileSize := 0
    }
    
    LogAction("Config file exists with size: " . fileSize . " bytes")
    if (fileSize <= 0) {
        LogAction("Config file exists but is empty")
        capturedPositions := Map()
        lastIconSequence := ""
        lastRepetitionCount := 1
        lastWaitTimeBetweenCycles := 5
        return
    }
    capturedPositions := Map()
    lastIconSequence := ""
    lastRepetitionCount := 1
    lastWaitTimeBetweenCycles := 5
    
    try {
        fileContent := FileRead(configFile)
        LogAction("Successfully read config file content with length: " . StrLen(fileContent))
        lines := StrSplit(fileContent, "`n", "`r")
        lineCount := lines.Length
        LogAction("Config file contains " . lineCount . " lines")
        positionsFound := 0
        
        Loop lineCount {
            line := lines[A_Index]
            
            ; Check for section headers
            if (line == "[Positions]") {
                inPositionsSection := true
                inBatchSettingsSection := false
                continue
            } else if (line == "[BatchSettings]") {
                inPositionsSection := false
                inBatchSettingsSection := true
                continue
            }
            
            ; Skip empty lines
            if (line == "")
                continue
            
            ; Process positions
            if (inPositionsSection) {
                pos := InStr(line, "=")
                if (pos > 0) {
                    key := SubStr(line, 1, pos-1)
                    value := SubStr(line, pos+1)
                    coords := StrSplit(value, ",")
                    if (coords.Length == 2) {
                        if (!capturedPositions.Has(key))
                            capturedPositions[key] := Map()
                        capturedPositions[key]["X"] := coords[1]
                        capturedPositions[key]["Y"] := coords[2]
                        positionsFound++
                    }
                }
            }
            
            ; Process batch settings
            if (inBatchSettingsSection) {
                pos := InStr(line, "=")
                if (pos > 0) {
                    key := SubStr(line, 1, pos-1)
                    value := SubStr(line, pos+1)
                    
                    if (key == "IconSequence") {
                        lastIconSequence := value
                        LogAction("Loaded IconSequence: " . value)
                    } else if (key == "RepetitionCount") {
                        lastRepetitionCount := Integer(value)
                        LogAction("Loaded RepetitionCount: " . value)
                    } else if (key == "WaitTime") {
                        lastWaitTimeBetweenCycles := Float(value)
                        LogAction("Loaded WaitTime: " . value)
                    }
                }
            }
        }
        
        LogAction("Loaded " . positionsFound . " positions from config file")
        if (positionsFound == 0 && lineCount > 1) {
            LogAction("WARNING: File has content but no valid positions were found")
            MsgBox("Warning: The config file exists but no valid positions were found.")
        }
    } catch {
        LogAction("ERROR: Exception when reading config file")
        MsgBox("Error reading config file. Check the log for details.")
    }
    
    ; Log first loaded position for debugging
    if (capturedPositions.Count > 0) {
        for key, pos in capturedPositions {
            LogAction("First position found: " . key . " = " . pos["X"] . "," . pos["Y"])
            break
        }
    }
}


; -----------------------------------------------------------------------------
; Ensures that the game window is active and resized.
EnsureGameWindowActive() {
    global appPath, windowWidth, windowHeight
    Local WinX, WinY, WinW, WinH
    if WinExist("ahk_exe LegendsOfIdleon.exe") {
        WinActivate("ahk_exe LegendsOfIdleon.exe")
        LogAction("Activated existing Idleon window")
        Sleep(1000)
    } else {
        LogAction("Launching Idleon game")
        Run(appPath)
        try {
            WinWait("ahk_exe LegendsOfIdleon.exe", , 30)
        } catch {
            MsgBox("Failed to start Idleon.")
            return false
        }
        Sleep(5000)
    }
    ResizeGameWindow()  ; resize the window
    return true
}

; -----------------------------------------------------------------------------
; Ensures that the alchemy screen is open using the new FindImageOnScreen function.
EnsureAlchemyScreenOpen(WinW, WinH) {
    Local alchCheckPath, alchCheckX, alchCheckY, alchemyButtonPath, alchemyBtnX, alchemyBtnY, alchemyBtnWidth, alchemyBtnHeight, alchemyBtnCenterX, alchemyBtnCenterY
    
    alchCheckPath := A_ScriptDir . "\alch_check.png"
    if (!FileExist(alchCheckPath)) {
        LogAction("alch_check.png not found at: " . alchCheckPath)
        MsgBox("Check image file (alch_check.png) not found! Aborting.")
        return false
    }
    
    ; Check if alchemy screen is already open using the new function
    if (FindImageOnScreen(alchCheckPath, &alchCheckX, &alchCheckY)) {
        LogAction("Alchemy window already open")
        return true  ; Alchemy screen is already open
    }
    
    LogAction("Alchemy screen not detected. Attempting to open it.")
    if (!WinActive("ahk_exe LegendsOfIdleon.exe")) {
        WinActivate("ahk_exe LegendsOfIdleon.exe")
        Sleep(1000)
    }
    LogAction("Pressing Escape to close any open dialogs")
    Send("{Escape}")
    Sleep(1000)
    LogAction("Opening codex")
    Send("c")
    Sleep(1500)
    
    alchemyButtonPath := A_ScriptDir . "\alchemy_button.png"
    if (!FileExist(alchemyButtonPath)) {
        LogAction("alchemy_button.png not found at: " . alchemyButtonPath)
        MsgBox("Alchemy button image file not found! Aborting.")
        return false
    }
    
    ; Find the alchemy button using the new function
    if (!FindImageOnScreen(alchemyButtonPath, &alchemyBtnX, &alchemyBtnY, &alchemyBtnWidth, &alchemyBtnHeight, "", , true, "alchemy_btn")) {
        LogAction("Alchemy button not found. Cannot open alchemy screen.")
        MsgBox("Alchemy button not found. Aborting.")
        return false
    }
    
    alchemyBtnCenterX := alchemyBtnX + (alchemyBtnWidth / 2)
    alchemyBtnCenterY := alchemyBtnY + (alchemyBtnHeight / 2)
    LogAction("Opening Alchemy page")
    MouseClick("left", alchemyBtnCenterX, alchemyBtnCenterY)
    Sleep(1500)
    
    ; Verify alchemy screen is now open
    if (FindImageOnScreen(alchCheckPath, &alchCheckX, &alchCheckY)) {
        LogAction("Alchemy screen successfully opened.")
        return true
    } else {
        LogAction("Failed to open alchemy screen after clicking alchemy button.")
        MsgBox("Failed to open alchemy screen. Ensure the alchemy button is visible.")
        return false
    }
}

; -----------------------------------------------------------------------------
; Initializes automated position capture for all sets, pages, and icons.
StartAutomatedPositionCapture() {
    global autoCapturing, captureQueue, currentSet, currentPage, currentTarget, currentIcon, isPositionFinderActive
    Local setNum, pageNum, iconNum, queueLength
    captureQueue := []            ; Clear previous queue
    autoCapturing := true
    isPositionFinderActive := true

    ; Queue navigation arrows and icon positions for 4 sets, 6 pages, 5 icons per page
    Loop 4 {
        setNum := A_Index
        captureQueue.Push([setNum, 0, "DownArrow"])
        captureQueue.Push([setNum, 0, "UpArrow"])
        Loop 6 {
            pageNum := A_Index
            Loop 5 {
                iconNum := A_Index
                captureQueue.Push([setNum, pageNum, "Icon", iconNum])
            }
        }
    }

    if (captureQueue.Length > 0) {
        SetTimer(UpdatePositionDisplay, 100)
        firstPosition := captureQueue.RemoveAt(1)
        currentSet := firstPosition[1]
        currentPage := firstPosition[2]
        currentTarget := firstPosition[3]
        if (currentTarget == "Icon")
            currentIcon := firstPosition[4]
        queueLength := captureQueue.Length + 1  ; include the current one
        MsgBox("Starting automated position capture for all sets and icons!`n`nThere are " . queueLength . " positions to capture.`n1. Move your mouse to the indicated position.`n2. Press Mouse5 to capture.`n3. Repeat until all positions are captured.`n`nPress OK to begin.", "Automated Position Capture")
        
        tooltipText := "Position: Set " . currentSet
        if (currentTarget == "DownArrow" || currentTarget == "UpArrow")
            tooltipText .= " " . currentTarget
        else
            tooltipText .= " Page " . currentPage . ", Icon " . currentIcon
        tooltipText .= "`nPress Mouse5 when ready"
        ToolTip(tooltipText)
        LogAction("Starting automated capture with " . captureQueue.Length . " remaining positions")
    }
}

; -----------------------------------------------------------------------------
; Saves captured positions and batch settings to the configuration file.
SavePositionsToFile() {
    global capturedPositions, configFile, lastIconSequence, lastRepetitionCount, lastWaitTimeBetweenCycles
    Local posCount, fileContent, fileSize, key, pos
    posCount := capturedPositions.Count
    LogAction("Attempting to save " . posCount . " positions to file: " . configFile)
    
    if (FileExist(configFile)) {
        FileCopy(configFile, configFile . ".bak", 1)
        LogAction("Backup created: " . configFile . ".bak")
    }
    
    fileContent := "[Positions]`n"
    for key, pos in capturedPositions {
        fileContent .= key . "=" . pos["X"] . "," . pos["Y"] . "`n"
    }
    
    ; Add batch settings section
    fileContent .= "`n[BatchSettings]`n"
    fileContent .= "IconSequence=" . lastIconSequence . "`n"
    fileContent .= "RepetitionCount=" . lastRepetitionCount . "`n"
    fileContent .= "WaitTime=" . lastWaitTimeBetweenCycles . "`n"
    
    try {
        FileDelete(configFile)
        LogAction("Deleted existing config file if present")
        FileAppend(fileContent, configFile)
        if (FileExist(configFile)) {
            fileSize := FileGetSize(configFile)
            if (fileSize > 0)
                LogAction("SUCCESS: Saved " . posCount . " positions and batch settings to " . configFile . " (" . fileSize . " bytes)")
            else {
                LogAction("ERROR: File created but empty!")
                MsgBox("Error: Config file was created but is empty! Check permissions.")
            }
        } else {
            LogAction("ERROR: Failed to create config file!")
            MsgBox("Error: Failed to create config file! Check path and permissions.")
        }
    } catch {
        LogAction("ERROR: Exception when saving file")
        MsgBox("Error saving positions. Check log for details.")
    }
    
    SplitPath(configFile, &fileName, &fileDir)
}

; -----------------------------------------------------------------------------
; Resizes the game window to the configured dimensions.
ResizeGameWindow() {
    global windowWidth, windowHeight
    Local WinX, WinY, WinW, WinH
    LogAction("Resizing window to " . windowWidth . "x" . windowHeight)
    WinRestore("ahk_exe LegendsOfIdleon.exe")
    WinMove(, , windowWidth, windowHeight, "ahk_exe LegendsOfIdleon.exe")
    Sleep(1000)
    WinGetPos(&WinX, &WinY, &WinW, &WinH, "ahk_exe LegendsOfIdleon.exe")
    LogAction("Window resized to: X=" . WinX . ", Y=" . WinY . ", W=" . WinW . ", H=" . WinH)
}

; -----------------------------------------------------------------------------
; Improved version of ProcessSpecificIcon using the new FindImageOnScreen function
ProcessSpecificIcon(setNum, pageNum, iconNum) {
    global capturedPositions, buttonImageFile, windowWidth, windowHeight, batchProcessingActive
    Local WinX, WinY, WinW, WinH, downArrowKey, upArrowKey, downArrowX, downArrowY, upArrowX, upArrowY
    Local iconKey, iconX, iconY, buttonFile, buttonX, buttonY, buttonWidth, buttonHeight, buttonCenterX, buttonCenterY

    ; Ensure window is active and sized
    EnsureGameWindowActive()
    WinGetPos(&WinX, &WinY, &WinW, &WinH, "ahk_exe LegendsOfIdleon.exe")
    if (WinW != windowWidth || WinH != windowHeight) {
        LogAction("Window size mismatch, resizing")
        ResizeGameWindow()
        WinGetPos(&WinX, &WinY, &WinW, &WinH, "ahk_exe LegendsOfIdleon.exe")
    }
    if (!WinActive("ahk_exe LegendsOfIdleon.exe")) {
        WinActivate("ahk_exe LegendsOfIdleon.exe")
        Sleep(1000)
    }
    ; Ensure alchemy screen is open
    if (!EnsureAlchemyScreenOpen(WinW, WinH))
        return false

    LogAction("Processing Set " . setNum . ", Page " . pageNum . ", Icon " . iconNum)

    ; Validate navigation arrow positions
    downArrowKey := "Set" . setNum . "_DownArrow"
    upArrowKey := "Set" . setNum . "_UpArrow"
    if (!capturedPositions.Has(downArrowKey) || !capturedPositions.Has(upArrowKey)) {
        LogAction("Missing navigation arrow positions for Set " . setNum)
        MsgBox("Missing navigation arrow positions for Set " . setNum . ". Please use position finder to set them.")
        return false
    }
    downArrowX := capturedPositions[downArrowKey]["X"]
    downArrowY := capturedPositions[downArrowKey]["Y"]
    upArrowX := capturedPositions[upArrowKey]["X"]
    upArrowY := capturedPositions[upArrowKey]["Y"]

    ; Reset to page 1 by clicking the down arrow several times
    LogAction("Resetting to page 1 by clicking down arrow 7 times")
    Loop 7 {
        MouseClick("left", downArrowX, downArrowY)
        Sleep(300)
    }
    ; Navigate to desired page using up arrow
    if (pageNum > 1) {
        navigateClicks := pageNum - 1
        LogAction("Navigating to page " . pageNum . " by clicking up arrow " . navigateClicks . " times")
        Loop navigateClicks {
            MouseClick("left", upArrowX, upArrowY)
            Sleep(300)
        }
    }
    ; Process the specified icon
    iconKey := "Set" . setNum . "_Page" . pageNum . "_Icon" . iconNum
    if (!capturedPositions.Has(iconKey)) {
        LogAction("Position for " . iconKey . " not found")
        MsgBox("Position for " . iconKey . " not found. Please use position finder to set it.")
        return false
    }
    iconX := capturedPositions[iconKey]["X"]
    iconY := capturedPositions[iconKey]["Y"]
    LogAction("Clicking icon " . iconNum . " at position: " . iconX . ", " . iconY)
    MouseClick("left", iconX, iconY)
    Sleep(1000)

    ; Click the upgrade button using our new image recognition function
    buttonFile := A_ScriptDir . "\" . buttonImageFile
    
    ; Use the new function to find the button
    debugPrefix := "debug_" . setNum . "_" . pageNum . "_" . iconNum
    if (FindImageOnScreen(buttonFile, &buttonX, &buttonY, &buttonWidth, &buttonHeight, 
                         "", , true, debugPrefix)) {
        
        buttonCenterX := buttonX + (buttonWidth / 2)
        buttonCenterY := buttonY + (buttonHeight / 2)
        
        LogAction("Found upgrade button at: " . buttonX . ", " . buttonY . " - Clicking center: " . buttonCenterX . ", " . buttonCenterY)
        MouseClick("left", buttonCenterX, buttonCenterY)
        Sleep(1000)
        
        ; Try to click the icon again to close the upgrade dialog
        LogAction("Clicking icon again at position: " . iconX . ", " . iconY)
        MouseClick("left", iconX, iconY)
        LogAction("Completed processing icon")
        return true
    } else {
        ; Don't show dialog in batch mode to avoid interrupting the process
        if (!batchProcessingActive) {
            MsgBox("Could not find the upgrade button on screen. A debug screenshot has been saved.")
        } else {
            LogAction("Continuing batch process despite upgrade button detection failure")
        }
        return false
    }
}

; -----------------------------------------------------------------------------
; Toggles the position finder mode for capturing UI element positions.
TogglePositionFinder() {
    global isPositionFinderActive, currentSet, currentPage, currentTarget, currentIcon
    Local setNum, pageNum, iconNum, elementType
    isPositionFinderActive := !isPositionFinderActive
    if (isPositionFinderActive) {
        SetTimer(UpdatePositionDisplay, 100)
        MsgBox("Position finder is now ACTIVE.`n`nHotkeys:`n- Mouse5 (XButton2): Capture current position`n- Ctrl+Shift+P: Toggle finder on/off`n- Ctrl+Shift+S: Save positions`n`nFollow the prompts:`n1. Set (1-4)`n2. Page (1-6 or 0 for arrows)`n3. Element (D for Down, U for Up, or Icon number 1-5)", "Position Finder Active")
        
        try {
            setNum := InputBox("Enter set number (1-4):", "Set Number", "w200 h130").Value
            if (setNum < 1 || setNum > 4) {
                isPositionFinderActive := false
                SetTimer(UpdatePositionDisplay, 0)
                return
            }
        } catch {
            isPositionFinderActive := false
            SetTimer(UpdatePositionDisplay, 0)
            return
        }
        
        try {
            pageNum := InputBox("Enter page number (1-6) or 0 for navigation arrows:", "Page Number", "w200 h130").Value
            if (pageNum < 0 || pageNum > 6) {
                isPositionFinderActive := false
                SetTimer(UpdatePositionDisplay, 0)
                return
            }
        } catch {
            isPositionFinderActive := false
            SetTimer(UpdatePositionDisplay, 0)
            return
        }
        
        if (pageNum == 0) {
            try {
                elementType := InputBox("Enter element type:`nD: Down Arrow`nU: Up Arrow", "Element Type", "w200 h170").Value
                if (elementType != "D" && elementType != "U") {
                    isPositionFinderActive := false
                    SetTimer(UpdatePositionDisplay, 0)
                    return
                }
            } catch {
                isPositionFinderActive := false
                SetTimer(UpdatePositionDisplay, 0)
                return
            }
            currentSet := setNum
            currentPage := 0
            currentTarget := (elementType == "D") ? "DownArrow" : "UpArrow"
        } else {
            try {
                iconNum := InputBox("Enter icon number (1-5):", "Icon Number", "w200 h130").Value
                if (iconNum < 1 || iconNum > 5) {
                    isPositionFinderActive := false
                    SetTimer(UpdatePositionDisplay, 0)
                    return
                }
            } catch {
                isPositionFinderActive := false
                SetTimer(UpdatePositionDisplay, 0)
                return
            }
            currentSet := setNum
            currentPage := pageNum
            currentIcon := iconNum
            currentTarget := "Icon"
        }
        LogAction("Finder activated for Set " . currentSet . ", Page " . currentPage . ", Target: " . currentTarget . (currentTarget == "Icon" ? " " . currentIcon : ""))
    } else {
        SetTimer(UpdatePositionDisplay, 0)
        ToolTip()
        LogAction("Finder deactivated")
    }
}

; -----------------------------------------------------------------------------
; Continuously displays the current mouse position relative to the game window.
UpdatePositionDisplay() {
    global isPositionFinderActive, currentSet, currentPage, currentTarget, currentIcon
    Local MouseX, MouseY, WindowID, WinTitle, WinClass, WinX, WinY, WinW, WinH, RelativeX, RelativeY, RelPercentX, RelPercentY, tooltipText
    if (!isPositionFinderActive)
        return
        
    MouseGetPos(&MouseX, &MouseY, &WindowID)
    WinTitle := WinGetTitle("ahk_id " . WindowID)
    WinClass := WinGetClass("ahk_id " . WindowID)
    
    if (InStr(WinClass, "Legends") || InStr(WinTitle, "Idleon")) {
        WinGetPos(&WinX, &WinY, &WinW, &WinH, "ahk_id " . WindowID)
        RelativeX := MouseX - WinX
        RelativeY := MouseY - WinY
        if (WinW > 0 && WinH > 0) {
            RelPercentX := Round(RelativeX / WinW * 100, 2)
            RelPercentY := Round(RelativeY / WinH * 100, 2)
            tooltipText := "Pos: " . RelativeX . ", " . RelativeY . " (window-relative)"
            tooltipText .= "`nRelative %: " . RelPercentX . "%, " . RelPercentY . "%"
            tooltipText .= "`nSet:" . currentSet . " Page:" . currentPage . " Target:" . currentTarget
            if (currentTarget == "Icon")
                tooltipText .= " " . currentIcon
            tooltipText .= "`nPress Mouse5 to capture"
            ToolTip(tooltipText)
        }
    } else {
        tooltipText := "Pos: " . MouseX . ", " . MouseY . " (screen coords - NOT OVER GAME)"
        tooltipText .= "`nSet:" . currentSet . " Page:" . currentPage . " Target:" . currentTarget
        if (currentTarget == "Icon")
            tooltipText .= " " . currentIcon
        tooltipText .= "`nMove mouse over game then press Mouse5"
        ToolTip(tooltipText)
    }
}

; -----------------------------------------------------------------------------
; Captures the current mouse position and stores it for the active UI element.
CaptureCurrentPosition() {
    global capturedPositions, currentSet, currentPage, currentTarget, currentIcon, autoCapturing, captureQueue
    Local RelativeX, RelativeY, key, queueLength, nextPosition, tooltipText
    MouseGetPos(&RelativeX, &RelativeY)
    if (currentTarget == "DownArrow" || currentTarget == "UpArrow")
        key := "Set" . currentSet . "_" . currentTarget
    else
        key := "Set" . currentSet . "_Page" . currentPage . "_Icon" . currentIcon
    if (!capturedPositions.Has(key))
        capturedPositions[key] := Map()
    capturedPositions[key]["X"] := RelativeX
    capturedPositions[key]["Y"] := RelativeY
    tooltipText := "Captured " . key . ": " . RelativeX . ", " . RelativeY . " (window-relative)"
    ToolTip(tooltipText)
    Sleep(1000)
    LogAction("Captured " . key . ": X=" . RelativeX . ", Y=" . RelativeY)
    queueLength := captureQueue.Length
    if (autoCapturing && queueLength > 0) {
        nextPosition := captureQueue.RemoveAt(1)
        currentSet := nextPosition[1]
        currentPage := nextPosition[2]
        currentTarget := nextPosition[3]
        if (currentTarget == "Icon")
            currentIcon := nextPosition[4]
        tooltipText := "Next: Set " . currentSet
        if (currentTarget == "DownArrow" || currentTarget == "UpArrow")
            tooltipText .= " " . currentTarget
        else
            tooltipText .= " Page " . currentPage . ", Icon " . currentIcon
        tooltipText .= "`nPress Mouse5 when ready"
        ToolTip(tooltipText)
        LogAction("Ready for next position: Set " . currentSet . ((currentTarget == "Icon") ? (" Icon " . currentIcon) : (" " . currentTarget)))
    } else if (autoCapturing) {
        MsgBox("All positions captured! Positions saved to file.")
        SavePositionsToFile()
        autoCapturing := false
        SetTimer(UpdatePositionDisplay, 0)
        ToolTip()
    } else {
        result := MsgBox("Do you want to capture another position?", "Continue Capturing?", "YesNo")
        if (result == "Yes") {
            try {
                setNum := InputBox("Enter set number (1-4):", "Set Number", "w200 h130").Value
                if (setNum < 1 || setNum > 4)
                    return
            } catch {
                return
            }
            
            try {
                pageNum := InputBox("Enter page number (1-6) or 0 for navigation arrows:", "Page Number", "w200 h130").Value
                if (pageNum < 0 || pageNum > 6)
                    return
            } catch {
                return
            }
            
            if (pageNum == 0) {
                try {
                    elementType := InputBox("Enter element type:`nD: Down Arrow`nU: Up Arrow", "Element Type", "w200 h170").Value
                    if (elementType != "D" && elementType != "U")
                        return
                } catch {
                    return
                }
                currentSet := setNum
                currentPage := 0
                currentTarget := (elementType == "D") ? "DownArrow" : "UpArrow"
            } else {
                try {
                    iconNum := InputBox("Enter icon number (1-5):", "Icon Number", "w200 h130").Value
                    if (iconNum < 1 || iconNum > 5)
                        return
                } catch {
                    return
                }
                currentSet := setNum
                currentPage := pageNum
                currentIcon := iconNum
                currentTarget := "Icon"
            }
            LogAction("Now capturing: Set " . currentSet . ", Page " . currentPage . ", Target: " . currentTarget . ((currentTarget == "Icon") ? (" " . currentIcon) : ("")))
        }
    }
}

; -----------------------------------------------------------------------------
; ------------------------ SCRIPT INITIALIZATION ----------------------------
#SingleInstance Force
SetWorkingDir A_ScriptDir
A_BatchLines := -1  ; Correct syntax for AHK v2 instead of SetBatchLines(-1)
CoordMode "Mouse", "Window"
CoordMode "Pixel", "Window"

; Define ExitFunc early
ExitFunc(ExitReason := "", ExitCode := "") {
    global pToken, batchProcessingActive
    if (batchProcessingActive)
        StopBatchProcessing()
    ; Note: We no longer shut down pToken here because it's reused
    LogAction("Script terminated and resources cleaned up")
}

; Include the GDI+ library
#Include Gdip_All.ahk

; ===== CONFIGURATION =====
global appPath := "E:\SteamLibrary\steamapps\common\Legends of Idleon\LegendsOfIdleon.exe"
global buttonImageFile := "upgrade_button.png"
global windowWidth := 1440
global windowHeight := 840
global configFile := A_ScriptDir . "\IdleonPositions.ini"

; ===== GLOBAL VARIABLES =====
global isPositionFinderActive := false
global capturedPositions := Map()
global currentTarget := ""
global currentSet := 0
global currentPage := 0
global currentIcon := 0
global autoCapturing := false
global captureQueue := []
global batchProcessingActive := false
global iconSequence := []
global repetitionCount := 1
global waitTimeBetweenCycles := 5  ; in minutes
global currentIconIndex := 1
global currentRepetition := 1
global scriptVersion := "1.0.0"  ; Current version of your script
global githubRepo := "YourUsername/YourRepoName"  ; Your GitHub repo
global updateCheckIntervalDays := 1  

; ===== INITIALIZATION =====
LogAction("Script started. Working directory set to: " . A_ScriptDir)
testFile := A_ScriptDir . "\test_write_permission.txt"
try {
    FileAppend("Test file for write permission`n", testFile)
    if (FileExist(testFile)) {
        LogAction("Successfully created test file; write permissions confirmed")
        FileDelete(testFile)
    } else {
        LogAction("CRITICAL ERROR: Cannot write files to script directory!")
        MsgBox("Cannot write files to the script directory!`nRun as administrator or change directory.", "Critical Error", 16)
    }
} catch {
    LogAction("CRITICAL ERROR: Cannot write files to script directory!")
    MsgBox("Cannot write files to the script directory!`nRun as administrator or change directory.", "Critical Error", 16)
}

; Global GDI+ initialization (pToken is reused)
global pToken := Gdip_Startup()
if (!pToken) {
    MsgBox("GDI+ failed to start. Please ensure GDI+ is installed.", "GDI+ Error", 48)
    ExitApp()
}
OnExit(ExitFunc)  ; Using function object directly
LoadPositionsFromFile()

; Launch the GUI automatically when script starts
StartBatchProcessing()

; -----------------------------------------------------------------------------
; ---- HOTKEYS ----
HotIfWinActive("ahk_exe LegendsOfIdleon.exe")
^+F1::StartBatchProcessing()  ; Start batch processing
^+F2::
{
    try {
        setNum := InputBox("Enter set number (1-4):", "Set Number", "w200 h130").Value
        pageNum := InputBox("Enter page number (1-6):", "Page Number", "w200 h130").Value
        iconNum := InputBox("Enter icon number (1-5):", "Icon Number", "w200 h130").Value
        
        if (setNum < 1 || setNum > 4 || pageNum < 1 || pageNum > 6 || iconNum < 1 || iconNum > 5) {
            MsgBox("Invalid input. Set: 1-4, Page: 1-6, Icon: 1-5")
            return
        }
        ProcessSpecificIcon(setNum, pageNum, iconNum)
    } catch {
        return
    }
}

^+p::TogglePositionFinder()    ; Toggle position finder mode
^!a::StartAutomatedPositionCapture()  ; Start automated position capture
XButton2::  ; Mouse5 to capture current position
{
    if (!isPositionFinderActive) {
        MsgBox("Position finder is not active. Press Ctrl+Shift+P first.")
        return
    }
    CaptureCurrentPosition()
}

^+s::  ; Save positions hotkey
{
    SavePositionsToFile()
    MsgBox("Positions saved to file: " . configFile)
}

; Add new hotkey for recapturing the upgrade button
^+u::BackupAndCaptureUpgradeButton()

^+x::  ; Stop batch processing and exit script
{
    if (batchProcessingActive) {
        StopBatchProcessing()
        MsgBox("Batch processing stopped. Exiting script.")
    } else {
        MsgBox("No batch processing is currently active. Exiting script.")
    }
    ExitApp()  ; This will terminate the script completely
}
HotIf

; -----------------------------------------------------------------------------
; Creates a GUI to configure batch processing parameters, populated with last used values.
StartBatchProcessing() {
    global appPath, BatchGui, lastIconSequence, lastRepetitionCount, lastWaitTimeBetweenCycles
    ; Ensure the game is running, active, and resized
    if (!EnsureGameWindowActive())
        return
    
    ; Create the Batch Processing Setup GUI
    BatchGui := Gui("", "Batch Processing Setup")
    BatchGui.Add("Text", "", "Bubble Sequence (format: [set,page,icon],[set,page,icon],...):")
    BatchGui.Add("Edit", "vIconSequenceInput w400", lastIconSequence)
    BatchGui.Add("Text", "", "How many times to repeat?:")
    BatchGui.Add("Edit", "vRepetitionCountInput w100", lastRepetitionCount)
    BatchGui.Add("Text", "", "Wait Time Between repetitions (minutes, can be 0):")
    BatchGui.Add("Edit", "vWaitTimeInput w100", lastWaitTimeBetweenCycles)
    startBtn := BatchGui.Add("Button", "Default", "Start")
    startBtn.OnEvent("Click", BatchStart)
    cancelBtn := BatchGui.Add("Button", "", "Cancel")
    cancelBtn.OnEvent("Click", BatchCancel)
    BatchGui.Show()
}

BatchStart(ctrl, *) {
    global BatchGui, iconSequence, repetitionCount, waitTimeBetweenCycles, currentIconIndex, currentRepetition, batchProcessingActive
    global lastIconSequence, lastRepetitionCount, lastWaitTimeBetweenCycles
    iconCount := 0
    
    ; Get values from GUI
    saved := BatchGui.Submit(false)  ; false means don't hide the GUI
    
    if (saved.IconSequenceInput == "") {
        MsgBox("Please enter an icon sequence.")
        return
    }
    
    if (saved.RepetitionCountInput == "" || saved.RepetitionCountInput < 1 || saved.RepetitionCountInput != Round(saved.RepetitionCountInput)) {
        MsgBox("Please enter a valid positive integer for repetition count.")
        return
    }
    
    if (saved.WaitTimeInput == "" || saved.WaitTimeInput < 0 || (saved.WaitTimeInput + 0 != saved.WaitTimeInput)) {
        MsgBox("Please enter a valid non-negative number for wait time.")
        return
    }
    
    if (!ParseIconSequence(saved.IconSequenceInput)) {
        MsgBox("Invalid icon sequence format. Please use: [set,page,icon],[set,page,icon],...")
        return
    }
    
    ; Save the current settings for future use
    lastIconSequence := saved.IconSequenceInput
    lastRepetitionCount := saved.RepetitionCountInput
    lastWaitTimeBetweenCycles := saved.WaitTimeInput
    SavePositionsToFile()  ; Save to config file
    
    repetitionCount := saved.RepetitionCountInput
    waitTimeBetweenCycles := saved.WaitTimeInput
    currentIconIndex := 1
    currentRepetition := 1
    batchProcessingActive := true
    iconCount := iconSequence.Length
    LogAction("Starting batch processing with " . iconCount . " icons, " . repetitionCount . " repetitions, and " . waitTimeBetweenCycles . " minute wait time")
    BatchGui.Destroy()
    
    MsgBox("Started batch processing with:`n- " . repetitionCount . " repetitions`n- " . waitTimeBetweenCycles . " minute wait between cycles`n- Processing " . iconCount . " icons in sequence`n`nPress Ctrl+Shift+X to stop at any time.", "Batch Processing Started")
    ProcessNextIcon()
}

BatchCancel(ctrl, *) {
    global BatchGui
    BatchGui.Destroy()
}

; -----------------------------------------------------------------------------
; Parses the icon sequence input and populates the global iconSequence array.
ParseIconSequence(input) {
    global iconSequence
    iconSequence := []
    
    input := Trim(input)
    bracketLevel := 0
    currentPart := ""
    
    Loop Parse, input {
        char := A_LoopField
        if (char = "[") {
            bracketLevel++
            currentPart .= char
        } else if (char = "]") {
            bracketLevel--
            currentPart .= char
            if (bracketLevel = 0) {
                pattern := "^\[(\d+),(\d+),(\d+)\]$"
                if (RegExMatch(currentPart, pattern, &match)) {
                    set := match[1]
                    page := match[2]
                    iconVal := match[3]
                    
                    if (set < 1 || set > 4 || page < 1 || page > 6 || iconVal < 1 || iconVal > 5) {
                        LogAction("Invalid icon values: " . currentPart . " - Values out of range")
                        return false
                    }
                    
                    iconSequence.Push([set, page, iconVal])
                } else {
                    LogAction("Invalid icon format: " . currentPart)
                    return false
                }
                currentPart := ""
            }
        } else if (char = "," && bracketLevel = 0) {
            continue
        } else {
            currentPart .= char
        }
    }
    
    if (iconSequence.Length = 0) {
        LogAction("No valid icons found in sequence")
        return false
    }
    
    seqStr := ""
    for i, icon in iconSequence {
        seqStr .= "[" . icon[1] . "," . icon[2] . "," . icon[3] . "],"
    }
    seqStr := SubStr(seqStr, 1, StrLen(seqStr)-1)
    LogAction("Parsed icon sequence: " . seqStr)
    return true
}

; -----------------------------------------------------------------------------
; Processes the next icon in the batch sequence.
ProcessNextIcon() {
    global iconSequence, currentIconIndex, currentRepetition, repetitionCount, waitTimeBetweenCycles, batchProcessingActive
    if (!batchProcessingActive) {
        LogAction("Batch processing was stopped before completion")
        return
    }
    
    if (currentRepetition > repetitionCount) {
        LogAction("Completed all " . repetitionCount . " repetitions of the icon sequence")
        MsgBox("Batch processing completed successfully!")
        batchProcessingActive := false
        return
    }
    
    currentIcon := iconSequence[currentIconIndex]
    setNum := currentIcon[1]
    pageNum := currentIcon[2]
    iconNum := currentIcon[3]
    
    LogAction("Processing icon #" . currentIconIndex . "/" . iconSequence.Length 
        . " (Set " . setNum . ", Page " . pageNum . ", Icon " . iconNum . ") - Repetition " . currentRepetition . "/" . repetitionCount)
        
    success := ProcessSpecificIcon(setNum, pageNum, iconNum)
    currentIconIndex++
    
    if (currentIconIndex > iconSequence.Length) {
        LogAction("Completed repetition " . currentRepetition . " of " . repetitionCount)
        if (currentRepetition < repetitionCount) {
            currentRepetition++
            currentIconIndex := 1
            waitTimeMs := waitTimeBetweenCycles * 60 * 1000
            if (waitTimeMs > 0) {
                LogAction("Waiting " . waitTimeBetweenCycles . " minutes before next repetition")
                ShowWaitingTimer(waitTimeMs, "Next repetition")
            } else {
                SetTimer(ProcessNextIcon, -100)
            }
        } else {
            LogAction("Completed all " . repetitionCount . " repetitions of the icon sequence")
            MsgBox("Batch processing completed successfully!")
            batchProcessingActive := false
            ExitApp()
        }
    } else {
        SetTimer(ProcessNextIcon, -500)
    }
}

; -----------------------------------------------------------------------------
; Starts a timer to display the waiting period between repetitions.
ShowWaitingTimer(waitTimeMs, action) {
    global waitTimerEndTime, waitTimerAction, batchProcessingActive
    waitTimerEndTime := A_TickCount + waitTimeMs
    waitTimerAction := action
    SetTimer(UpdateWaitingTimer, 1000)
}

; Updates the waiting timer display.
UpdateWaitingTimer() {
    global waitTimerEndTime, waitTimerAction, batchProcessingActive
    if (!batchProcessingActive) {
        SetTimer(UpdateWaitingTimer, 0)
        ToolTip()
        return
    }
    
    remaining := waitTimerEndTime - A_TickCount
    if (remaining <= 0) {
        SetTimer(UpdateWaitingTimer, 0)
        ToolTip()
        SetTimer(ProcessNextIcon, -100)
        return
    }
    
    remainingMinutes := Floor(remaining / 60000)
    remainingSeconds := Floor((remaining - (remainingMinutes * 60000)) / 1000)
    tooltipText := waitTimerAction . " in: " . remainingMinutes . "m " . remainingSeconds . "s`nPress Ctrl+Shift+X to cancel"
    ToolTip(tooltipText, 100, 100)
}

; -----------------------------------------------------------------------------
; Stops batch processing and clears any active timers/tooltips.
StopBatchProcessing() {
    global batchProcessingActive
    if (batchProcessingActive) {
        batchProcessingActive := false
        SetTimer(UpdateWaitingTimer, 0)
        SetTimer(ProcessNextIcon, 0)
        ToolTip()
        LogAction("Batch processing stopped by user")
    }
}

; -----------------------------------------------------------------------------
; Finds an image on screen using multiple tolerance values
; Returns true if found, and sets the position variables
; Parameters:
;   imagePath - path to the image file
;   &imageX, &imageY - variables to store the found coordinates
;   &width, &height - variables to store the image dimensions (optional)
;   searchRegion - search coordinates in format "x|y|w|h" (optional, defaults to entire window)
;   toleranceValues - array of tolerance values to try (optional)
;   saveDebugScreenshot - whether to save a debug screenshot if not found (optional)
;   debugPrefix - prefix for debug screenshot filename (optional)
FindImageOnScreen(imagePath, &imageX, &imageY, &width := 0, &height := 0, searchRegion := "", toleranceValues := [40, 80, 120, 150], saveDebugScreenshot := false, debugPrefix := "debug") {
    Local WinX, WinY, WinW, WinH, searchX, searchY, searchW, searchH, currentTolerance, debugScreenshot
    
    if (!FileExist(imagePath)) {
        LogAction("Image file not found: " . imagePath)
        return false
    }
    
    ; Get window dimensions if search region not specified
    if (searchRegion == "") {
        WinGetPos(&WinX, &WinY, &WinW, &WinH, "ahk_exe LegendsOfIdleon.exe")
        searchX := 0
        searchY := 0
        searchW := WinW
        searchH := WinH
    } else {
        searchParams := StrSplit(searchRegion, "|")
        if (searchParams.Length >= 4) {
            searchX := searchParams[1]
            searchY := searchParams[2]
            searchW := searchParams[3]
            searchH := searchParams[4]
        } else {
            WinGetPos(&WinX, &WinY, &WinW, &WinH, "ahk_exe LegendsOfIdleon.exe")
            searchX := 0
            searchY := 0
            searchW := WinW
            searchH := WinH
        }
    }
    
    ; Get image dimensions if requested
    if (&width != 0 && &height != 0) {
        if (!ImageGetSize(imagePath, &width, &height)) {
            width := 100  ; Default fallback values
            height := 30
            LogAction("Warning: Could not get image dimensions for " . imagePath . ", using defaults")
        }
    }
    
    LogAction("Searching for image: " . imagePath)
    
    ; Try multiple tolerance values
    for i, currentTolerance in toleranceValues {
        LogAction("Trying image search with tolerance: " . currentTolerance)
        
        try {
            if ImageSearch(&imageX, &imageY, searchX, searchY, searchX + searchW, searchY + searchH, "*" . currentTolerance . " " . imagePath) {
                LogAction("Found image with tolerance " . currentTolerance . " at " . imageX . ", " . imageY)
                return true
            }
        } catch {
            LogAction("Error during image search with tolerance " . currentTolerance)
        }
        
        ; Add a small delay between attempts
        Sleep(200)
    }
    
    ; If image not found and debug screenshot requested
    if (saveDebugScreenshot) {
        debugScreenshot := A_ScriptDir . "\" . debugPrefix . "_screenshot_" . FormatTime(, "yyyyMMdd_HHmmss") . ".png"
        SaveScreenshot(debugScreenshot)
        LogAction("Image not found. Debug screenshot saved to: " . debugScreenshot)
    }
    
    return false
}