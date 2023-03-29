#SingleInstance Force
SetWorkingDir %A_ScriptDir%
#InstallMouseHook
#Include FindText.ahk
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
Gui Add, Button, x65 y70 w100 h30 gSave_Position, Save Position Ctrl+Shft+S
Gui Add, Button, x175 y70 w100 h30 gLoad_Positions_From_File, Load Position(s) Ctrl+Shft+L
Gui Add, Button, x285 y70 w100 h30 gSet_To_Triage, Set to Triage Ctrl+Shft+T

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
    
    MsgBox, 4099, , Save by window title? (Yes) or by exe name? (No)
     WinActivate, %ActiveWinTitle%
    IfMsgBox, Yes
    {
        GuiControlGet, sel_title, , WindowTitle
        WinGetPos, Win_X, Win_Y, Win_W, Win_H, %sel_title%
        IniWrite, %Win_X%, saved_positions.txt, %sel_title%_%MyTab%, X
        IniWrite, %Win_Y%, saved_positions.txt, %sel_title%_%MyTab%, Y
        IniWrite, %Win_W%, saved_positions.txt, %sel_title%_%MyTab%, Width
        IniWrite, %Win_H%, saved_positions.txt, %sel_title%_%MyTab%, Height
        IniWrite, %MyTab%, saved_positions.txt, %sel_title%_%MyTab%, Tab
        Save_Type := "Window Title"
    }
    IfMsgBox, No
    {
        GuiControlGet, sel_title, , WindowTitle
        WinGetPos, Win_X, Win_Y, Win_W, Win_H, %sel_title%
        IniWrite, %Win_X%, saved_positions.txt, %this_exe%_%MyTab%, X
        IniWrite, %Win_Y%, saved_positions.txt, %this_exe%_%MyTab%, Y
        IniWrite, %Win_W%, saved_positions.txt, %this_exe%_%MyTab%, Width
        IniWrite, %Win_H%, saved_positions.txt, %this_exe%_%MyTab%, Height
        IniWrite, %MyTab%, saved_positions.txt, %this_exe%_%MyTab%, Tab
        Save_Type := "Exe Name"
    }
    IfMsgBox, Cancel
    {
        Return
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
          
            if (Tab == MyTab)
            {                
              ;  substring := SubStr(new_title, 1, 5)
              ;If WinExist(substring)
              ;  {
              ;      WinActivate
              ;      SetTitleMatchMode, 1 ;Use the 1 value to match the beginning of titles by the substring. 
              ;      WinMove, %substring%,, %Win_X%, %Win_Y%, %Win_W%, %Win_H%
              ;      
              ;      Sleep, 100
              ;  }

              exeFinal = ahk_exe %new_title%
              If WinExist(exeFinal)
                {
                    WinActivate %exeFinal%
                   ; SetTitleMatchMode, 1 ;Use the 1 value to match the beginning of titles by the substring. 
                    WinMove, %exeFinal%,, %Win_X%, %Win_Y%, %Win_W%, %Win_H%
                    
                    Sleep, 100
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
#IfWinActive, WindowMover.ahk ;ahk_class 
Delete::
ControlGetFocus, OutputVar, WindowMover.ahk
;if ErrorLevel
;    MsgBox, The target window doesn't exist or none of its controls has input focus.
;else
    

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

#IfWinActive
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

;===========FINDTEXT FUNCTIONS==========
; A function to read your messages with less clicking around.
^+r::
MouseClick, Right 
Sleep, 200


t1:=A_TickCount, Text:=X:=Y:=""

Text:="|<MarkRead>**50$68.kQ06000DU00S701U003M007Xk0M000n001chtqkwwBXbXv/3NM1s3lgBiKraQ3v0oTSFrB9b0aQBY4YRmmNMNX3BXBj8rqP3zUnDTTU"

if (ok:=FindText(X, Y, -523-150000, 690-150000, -523+150000, 690+150000, 0, 0, Text))
{
   FindText().Click(X, Y, "L")
}

Return

^+t::
Set_To_Triage:

t1:=A_TickCount, Text:=X:=Y:=""
Text:="|<TriageGray50>**50$32.zVU00100000E00004TjbrV2MDDAEaSGz48ggg12/BBAFaTTS0000U0003s2"

if (ok:=FindText(X, Y, -523-150000, 690-150000, -523+150000, 690+150000, 0, 0, Text))
{
   FindText().Click(X, Y, "L")
}
Sleep, 50

If WinExist("Tasks")
WinMinimize
Sleep, 50
Text:="|<MapGray50>**50$20.lU0AM03a00fjjecBCeSHagYt/9C2Ty00A0032"

if (ok:=FindText(X, Y, -1779-150000, 699-150000, -1779+150000, 699+150000, 0, 0, Text))
{
   FindText().Click(X, Y, "L")
}
Sleep, 50
If WinExist("Call Center")
WinMinimize
Text:="|<WikiGray50>**50$21.U7MS0V0n4M2NvTGd++J9lQd+NX9NAN9g"

if (ok:=FindText(X, Y, -1779-150000, 699-150000, -1779+150000, 699+150000, 0, 0, Text))
{
   FindText().Click(X, Y, "L")
}
Sleep, 1000
If WinExist("Wiki -")
WinMinimize
Gosub, Load_Positions_From_File
Return


        
;===========END OF FINDTEXT FUNCTIONS==========


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

; Remove empty lines or lines containing only spaces and > from clipboard (condenses emails)
; also removes a couple specific lines of text we don't need to record.
^#r::

Clipboard := StrReplace(Clipboard, "CAUTION: Not Open Dental internal. Do not click on untrusted items.")
Clipboard := RegExReplace(Clipboard, "m)(^\s*You don't often get email from \S+ Learn why this is important)", "")
Clipboard := RegExReplace(Clipboard, "m)(^[>\s]+\R)","")
return

; Email Sig... copies text from a file called "EmailSig.txt" and outputs it to the cursor
^+g::
IfNotExist, %A_ScriptDir%\EmailSig.txt
{
    FileAppend,, %A_ScriptDir%\EmailSig.txt
    MsgBox, 8240, DocWin - Email Signature needed, Email Signature file will open when you click OK. Add your name and role to it then save it. Example:`n`tChris D`n`tEmail and Fax Coordinator`n`nFile Name:`n%A_ScriptDir%\EmailSig.txt
    Run, %A_ScriptDir%\EmailSig.txt
}
FileRead, EmailSig, *t %A_ScriptDir%\EmailSig.txt
SendInput %EmailSig%
return

; patNumJump... puts "PatNum:" in front of an entered number so you can hit select to go right to the patient.
^+X::
ClipCheck := SubStr(RegExReplace(Clipboard, "[^0-9]", ""), 1, 10) ; get a 2-10 digit number from clipboard, ignoring all non-digit characters, to pre-populate the number.

Gui, PatNumJump:New, , Enter Pat/Task/Ph Num:
Gui, PatNumJump:Add, Edit, w120 vNumberCheck, %ClipCheck%
Gui, PatNumJump:Add, Button, vNoButton Default, &OK
GuiControl, hide, NoButton
Gui , -SysMenu
Gui, PatNumJump:Show, Autosize 
return

PatNumJumpButtonOK:
    Gui, Submit
    NumberCheck := RegExReplace(NumberCheck, "[^0-9]", "")
    if (StrLen(NumberCheck) < 2)
    {
        GoSub PatNumJumpGuiClose    
    }
    else if (StrLen(NumberCheck) <= 6)
    {
        clipboard := "PatNum:" . NumberCheck
    }
    else if (StrLen(NumberCheck) <= 9)
    {
        clipboard := "TaskNum:" . NumberCheck
    }
    else if (StrLen(NumberCheck) = 10)
    {
        clipboard := RegExReplace(NumberCheck, "(\d{3})(\d{3})(\d{4})", "($1) $2-$3")
    }

PatNumJumpGuiEscape:
PatNumJumpGuiClose:
    Gui, Destroy
    return
;=--
;=== End Useful Keys for All===



GuiClose:
ExitApp   