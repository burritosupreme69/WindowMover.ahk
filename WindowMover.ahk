#SingleInstance Force
SetWorkingDir %A_ScriptDir%
#InstallMouseHook
DetectHiddenWindows, On
start := A_TickCount
 if not A_IsAdmin
    Run *RunAs "%A_ScriptFullPath%" 



;============GUI AND WINDOW MOVING FUNCTIONS============
SetBatchLines, -1
WM_COMMAND := 0x0111
; Use descriptive variable names and capitalize keywords for clarity
Gui +AlwaysOnTop
Gui Margin, 0, 0

Gui Add, Text, x10 y10 w100 h20, Window Title:
Gui Add, Text, x120 y10 w200 h40 vWindowTitle
Gui Add, Text, x10 y40 w100 h20, Executable Name: 
Gui Add, Text, x120 y40 w200 h20 vEXE_Name, %EXE_Name%
; Add buttons with centered x values
Gui Add, Text, x10 y60 w100 h30 Center, Save Position Ctrl+Shft+S
Gui, Add, CheckBox, vSaveExeName gSaveExeName, Save as Exe Name(Notepad++ for instance)
Gui Add, Button, x120 y60 w100 h30 gLoad_Positions_From_File, Load Position(s) Ctrl+Shft+L
Gui Add, Button, x340 y60 w100 h30 gShow_Hotkeys, DocWin Hotkeys List
Gui Add, Tab3, x10 y110 gMyTab vMyTab, Profile 1|Profile 2|Profile 3

Gui Tab, 1
Gui Add, ListView, r10 w430 h200 gList1 vList1, Window Name
FileRead, data, saved_positions.txt
Loop, Parse, data, `n
{
    if RegExMatch(A_LoopField, "\[([^\]]+)\]", match) 
    {
        StringSplit, fields, match1, `t              
    }

    fields1 := RegExReplace(fields1, "_(Profile [123])$", "")
    if InStr(A_LoopField, "Tab=Profile 1")
    {
        LV_Add("", fields1)
    }
}

Gui Tab, 2
Gui Add, ListView, r10 w430 h200 gList2 vList2, Window Name
Loop, Parse, data, `n
{    
    if RegExMatch(A_LoopField, "\[([^\]]+)\]", match)
    {
        StringSplit, fields, match1, `t                             
    }
    
    fields1 := RegExReplace(fields1, "_(Profile [123])$", "")
    if InStr(A_LoopField, "Tab=Profile 2")
    {
        LV_Add("", fields1) 
    }     
}

Gui Tab, 3
Gui Add, ListView, r10 w430 h200 gList3 vList3, Window Name
Loop, Parse, data, `n
{
    if RegExMatch(A_LoopField, "\[([^\]]+)\]", match)
    {
        StringSplit, fields, match1, `t                          
    }

    fields1 := RegExReplace(fields1, "_(Profile [123])$", "")
    if InStr(A_LoopField, "Tab=Profile 3")
    {
        LV_Add("", fields1)
    }     
}

; Center the GUI on the screen
Gui, Show, w450 h350 NoActivate



WinGetActiveTitle, ActiveTitle
GuiControl, Text, WindowTitle, %ActiveTitle%
SetTimer, UpdateTitle, 200
return



;Insert the hotkey list here for quick viewing.
Show_Hotkeys:
Gui, HotkeysList: New, +AlwaysOnTop
Gui, HotkeysList: Add, Text, x10 y30 w90 h20, Ctrl+=
Gui, HotkeysList: Add, Text, x105 y30 w125 h20, Clear Clipboard
Gui, HotkeysList: Add, Text, x10 y60 w90 h20, Ctrl+/
Gui, HotkeysList: Add, Text, x105 y60 w125 h20, Start Calculator
Gui, HotkeysList: Add, Text, x10 y90 w90 h20, Ctrl+(Period)
Gui, HotkeysList: Add, Text, x105 y90 w125 h20, Open Snipping Tool
Gui, HotkeysList: Show, Autosize , Hotkeys List
Return



SaveExeName:
GuiControlGet, SaveExeName ; Retrieves 1 if it is checked, 0 if it is unchecked.
If (SaveExeName = 0)
{

return
}
else
{

return
}

UpdateTitle:

        WinGetActiveTitle, ActiveTitle
        GuiControl, Text, WindowTitle, %ActiveTitle%
        WinGetActiveStats, StatsTitle, Width, Height, X, Y
        WinGet, this_exe, ProcessName, %StatsTitle%
        WinGetPos, Win_X, Win_Y, Win_W, Win_H, %sel_title%
        GuiControl,, EXE_Name, %this_exe%
        GuiControl,, Window_Width, %Width%
        GuiControl,, Window_Height, %Height%
        GuiControl,, Window_X_Pos, %X%
        GuiControl,, Window_Y_Pos, %Y%
        MouseGetPos,x1,y1

Return

^+s::
Save_Position:
    Gosub, MyTab
    WinGetActiveTitle, ActiveWinTitle
    
    WinActivate, %ActiveWinTitle%
    if (SaveExeName = 1) {
        this_exe := SubStr(this_exe, 1, StrLen(this_exe)-4)
        StringUpper, this_exe, this_exe, T
        GuiControlGet, sel_title, , WindowTitle
        WinGetPos, Win_X, Win_Y, Win_W, Win_H, %sel_title%
        
        ; Keep only the first word in the string
        spacePos := InStr(this_exe, " ")
        underscorePos := InStr(this_exe, "_")
        
        ; Find the minimum position of space and underscore
        if (underscorePos > 0 && (spacePos == 0 || underscorePos < spacePos))
            spacePos := underscorePos
        
        if (spacePos > 0) {
            this_exe := SubStr(this_exe, 1, spacePos-1)
        }
        
        IniWrite, %Win_X%, saved_positions.txt, %this_exe%_%MyTab%, X
        IniWrite, %Win_Y%, saved_positions.txt, %this_exe%_%MyTab%, Y
        IniWrite, %Win_W%, saved_positions.txt, %this_exe%_%MyTab%, Width
        IniWrite, %Win_H%, saved_positions.txt, %this_exe%_%MyTab%, Height
        IniWrite, %MyTab%, saved_positions.txt, %this_exe%_%MyTab%, Tab
        IniWrite, "EXE", saved_positions.txt, %this_exe%_%MyTab%, SaveType ; Add this line to save the SaveType
        Save_Type := "Exe Name"
    }
    else {
        GuiControlGet, sel_title, , WindowTitle
        WinGetPos, Win_X, Win_Y, Win_W, Win_H, %sel_title%
        IniWrite, %Win_X%, saved_positions.txt, %sel_title%_%MyTab%, X
        IniWrite, %Win_Y%, saved_positions.txt, %sel_title%_%MyTab%, Y
        IniWrite, %Win_W%, saved_positions.txt, %sel_title%_%MyTab%, Width
        IniWrite, %Win_H%, saved_positions.txt, %sel_title%_%MyTab%, Height
        IniWrite, %MyTab%, saved_positions.txt, %sel_title%_%MyTab%, Tab
        IniWrite, "WT", saved_positions.txt, %sel_title%_%MyTab%, SaveType ; Add this line to save the SaveType
        Save_Type := "Window Title"
    }
    
    if (MyTab = "Profile 1")
    {
        if (Save_Type = "Window Title")
        {
            sel_title := sel_title
        }
        else
        {
            
            sel_title := this_exe
        }
        Gosub, List1
    }
    if (MyTab = "Profile 2")
    {
        if (Save_Type = "Window Title")
        {
            sel_title := sel_title 
        }
        else
        {
           
            sel_title := this_exe 
        }
        Gosub, List2
    }
    if (MyTab = "Profile 3")
    {
        if (Save_Type = "Window Title")
        {
            sel_title := sel_title
        }
        else
        {
          
            sel_title := this_exe
        }
        Gosub, List3
    }
Return

;Continue logic flow from previous function
^+l::
Load_Positions_From_File:
    Gosub, MyTab
    SetTimer, UpdateTitle, Off ; The timer being on can interrupt loading the window positions
    FileRead, data, saved_positions.txt
    Loop, Parse, data, `n
    {
        ; Extract data inside brackets using regex
        if RegExMatch(A_LoopField, "\[([^\]]+)\]", match)
        {
            ; Split extracted data into fields
            StringSplit, fields, match1, `t
            sel_title := fields1
            StringReplace, new_title, sel_title, _Profile 1, , All
            StringReplace, new_title, new_title, _Profile 2, , All
            StringReplace, new_title, new_title, _Profile 3, , All
            IniRead, Win_X, saved_positions.txt, %sel_title%, X
            IniRead, Win_Y, saved_positions.txt, %sel_title%, Y
            IniRead, Win_W, saved_positions.txt, %sel_title%, Width
            IniRead, Win_H, saved_positions.txt, %sel_title%, Height
            IniRead, Tab, saved_positions.txt, %sel_title%, Tab
            IniRead, SaveType, saved_positions.txt, %sel_title%, SaveType ; Add this line to read the SaveType
            
            if (Tab == MyTab)
            {                
                                    SetTitleMatchMode, 1 ;Use the 1 value to match the beginning of titles by the substring. 

                substring := SubStr(new_title, 1, 5)
                If WinExist(substring)
                {

                    WinActivate
                    WinMove, %substring%,, %Win_X%, %Win_Y%, %Win_W%, %Win_H%
                    Sleep, 100
                }
                
                SetTitleMatchMode, 2
                
                exeFinal = %new_title%
                
                If (SaveType = "EXE") ; Check the SaveType to determine if it's an exe name or window title
                {
                    If WinExist(exeFinal)
                    {      
                        WinActivate, %exeFinal%
                        WinMove, %exeFinal%,, %Win_X%, %Win_Y%, %Win_W%, %Win_H%
                        Sleep, 100
                    }
                }
            }
        }
    }
    SetTimer, UpdateTitle, On ;Restart the timer 
Return


;Updates the MyTab value which stores the current profile.
MyTab:
Gui, Submit, Nohide   
Return

;A glitch exists that will add a value containing the '_Profile X' value
List1:
Gosub, MyTab
If (MyTab = "Profile 1")
    {
        Gui, ListView, List1
        ; Loop through all items in the list and check if sel_title is already present
        Found := false
        Loop, % LV_GetCount()
        {
            LV_GetText(itemText, A_Index, 1)
            if (itemText = sel_title)
            {
                Found := true
                break
            }
        }
        ; If sel_title is not already present in the list, add it
        if (!Found)
            LV_Add("", sel_title)   
    }
Return



List2:
Gosub, MyTab
If (MyTab = "Profile 2")  
    {
        Gui, ListView, List2
        ; Loop through all items in the list and check if sel_title is already present
        Found := false
        Loop, % LV_GetCount()
        {
            LV_GetText(itemText, A_Index, 1)
            if (itemText = sel_title)
            {
                Found := true
                break
            }
        }
        ; If sel_title is not already present in the list, add it
        if (!Found)
            LV_Add("", sel_title)
    }
Return

List3:
Gosub, MyTab
If (MyTab = "Profile 3")
    {
        Gui, ListView, List3
        ; Loop through all items in the list and check if sel_title is already present
        Found := false
        Loop, % LV_GetCount()
        {
            LV_GetText(itemText, A_Index, 1)
            if (itemText = sel_title)
            {
                Found := true
                break
            }
        }
        ; If sel_title is not already present in the list, add it
        if (!Found)
            LV_Add("", sel_title)   
    }
Return

;change this to be a dynamic variable that will take in the scripts name
#If WinActive(A_ScriptName)

Delete::
ControlGetFocus, OutputVar, %A_ScriptName%

if (OutputVar = "SysListView321" or OutputVar = "SysListView322" or OutputVar = "SysListView323") {
   
    ; Add your code to execute when the hotkey is pressed
   ; MsgBox, Hit!
   Gosub, DeleteCheck
    return
}
else
{
    Return
}

#If

DeleteCheck:
Gosub, MyTab
If (MyTab = "Profile 1")  
    {
        Gui, ListView, List1
        ; Loop through all items in the list and check if sel_title is already present
        Found := false
        Loop, % LV_GetCount()
        {
            LV_GetText(itemText, A_Index, 1)
            if (itemText = sel_title)
            {
                Found := true
                break
            }
        }
            MsgBox, 4100, ,Delete Entry/Entries?
            IfMsgBox, Yes
            {
            LvSel := A_GuiControl
        ;LV_Delete(sel_title)
        sel_title = 
        Gosub, DeleteRow
            }
            else IfMsgBox, No
            {
                Return
            }
 
    }

If (MyTab = "Profile 2")  
    {
        Gui, ListView, List2
        ; Loop through all items in the list and check if sel_title is already present
        Found := false
        Loop, % LV_GetCount()
        {
            LV_GetText(itemText, A_Index, 1)
            if (itemText = sel_title)
            {
                Found := true
                break
            }
        }
 
            MsgBox, 4100, ,Delete Entry/Entries?
            IfMsgBox, Yes
            {
            LvSel := A_GuiControl
        ;LV_Delete(sel_title)
        sel_title = 
        Gosub, DeleteRow
            }
            else IfMsgBox, No
            {
                Return
            }
 
    }

If (MyTab = "Profile 3")  
    {
        Gui, ListView, List3
        ; Loop through all items in the list and check if sel_title is already present
        Found := false
        Loop, % LV_GetCount()
        {
            LV_GetText(itemText, A_Index, 1)
            if (itemText = sel_title)
            {
                Found := true
                break
            }
        }
        MsgBox, 4100, ,Delete Entry/Entries?
            IfMsgBox, Yes
            {
            LvSel := A_GuiControl
        ;LV_Delete(sel_title)
        sel_title = 
        Gosub, DeleteRow
            }
            else IfMsgBox, No
            {
                Return
            }
 
    }

Return


;Searching for each tab this way seems a little dense for coding, I'm sure this can be condensed.
DeleteRow: 
    RowNumber := 0, SelectedArray := []

    ; Get a list of selected items
    gui, ListView, %LvSel%
    Loop
    {
        RowNumber := LV_GetNext(RowNumber)
        if not RowNumber
            Break
        SelectedArray.Push(RowNumber)
    }

    ; Iterate through selected items in reverse order
 Loop, % SelectedArray.Length()
    {
        index := SelectedArray.Length() - A_Index + 1
        RowNumber := SelectedArray[index]
        LV_GetText(SelectedRowText, RowNumber)

        if (MyTab = "Profile 1") 
        {
            IniRead, data, saved_positions.txt
            Loop, Parse, data, `n
            {
                FinalDelete = %SelectedRowText%_Profile 1
                IniDelete, saved_positions.txt,%FinalDelete%                    
            }
        }
        else if (MyTab = "Profile 2") 
        {
            IniRead, data, saved_positions.txt
            Loop, Parse, data, `n
            {
                FinalDelete = %SelectedRowText%_Profile 2
                IniDelete, saved_positions.txt,%FinalDelete%                    
            }
        }
        else if (MyTab = "Profile 3") 
        {
            IniRead, data, saved_positions.txt
            Loop, Parse, data, `n
            {
                FinalDelete = %SelectedRowText%_Profile 3
                IniDelete, saved_positions.txt,%FinalDelete%                    
            }
        }

        ; Delete the row from ListView after processing
        LV_Delete(RowNumber)
    }

Return

;============END OF GUI AND WINDOW MOVING FUNCTIONS============


;--= Useful Keys for All ===
; Get current Customer # from customer's database
^0::
processName := "opendental.exe"
if (WinExist("ahk_exe " . processName))
{
    SetTitleMatchMode RegEx
    WinGetTitle, title, ^Customers.+
    splarr := StrSplit(title, "-")
    custnum := splarr[splarr.MaxIndex()]
    splarr := StrSplit(custnum, " ")
    custnum := splarr[2]
    SendInput %custnum%
}
return

; Clear Clipboard
^=::
Clipboard :=
return

; start Calculator
^/::
Run calc.exe
return

; open Snipping Tool
^.::
Run snippingtool.exe
return




GuiClose:
ExitApp   


