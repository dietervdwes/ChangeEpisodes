;Change Episodes 
;Written by Dieter van der Westhuizen 21-04-2020
;dietervdwes@gmail.com
;dieter.vdwesthuizen@nhls.ac.za

#SingleInstance, force
;#NoTrayIcon
#NoEnv
#Persistent
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetKeyDelay, 150
SendMode, Input
SetTitleMatchMode, 2
SetDefaultMouseSpeed, 0
SetMouseDelay, 0
SetWinDelay, 500



;;;;;;;;;;;;;;;;;;;;;;;;;;;;    Add Buttons ;;;;;;;;;;;;;;;;;;
Gui, Add, button, x2 y2 w75 h40 ,Change_Location
Gui, Add, button, x2 y44 w75 h20 ,Re_bill
Gui, Add, button, x2 y66 w75 h20 ,Clear_List
Gui, Add, button, x2 y88 w75 h20 ,Change_Ref
Gui, Add, button, x2 y110 w75 h20 ,Remove_fol
Gui, Add, button, x2 y132 w75 h20 ,Re-add_tests
Gui, Add, button, x2 y152 w32 h20 ,Close
Gui, Add, button, x45 y152 w20 h20 ,_i
;;;;;;;;;;;;;;;;;;;;;;;;;;;;   Set Window Options   ;;;;;;;;;;;;;;
;Gui, +AlwaysOnTop
Gui, -sysmenu +AlwaysOnTop
Gui, Show, , Change-episodes
WinGetPos,,,,TrayHeight,ahk_class Shell_TrayWnd,,,
height := A_ScreenHeight-270
width := A_ScreenWidth-85
Gui, Margin, 0, 0
;Gui, Add, Picture, x0 y0 w411 h485, picture.png
;Gui -Caption -Border
Gui, Show, x%width% y%height% w80

return

; This is to read the CSV defined in FileName below, parse the data and send to TrakCare window, also specified below, alter those episodes' details as specified.

ButtonChange_Location:
sleep, 500
IfWinNotExist, Patient Entry
        {
            MsgBox, Please open the Patient Entry Window
            Return
        }
WinActivate, Patient Entry

FileName := loc_ward_list.csv
Loop, read, loc_ward_list.csv, ;output_list.txt ; output_list.txt is the file to write to, if necessary
{   ;This part of code reads and parses the CSV / txt file and stores results of the read file in an array.
    LineNumber := A_Index
    Loop, parse, A_LoopReadLine, CSV
    
        EpisodeLine := LineNumber
        FileReadLine, Line_out, loc_ward_list.csv, EpisodeLine
        LineArray := StrSplit(Line_out, ",")
        Episode_out := LineArray[1]
        Location_code := LineArray[2] ; Not actually used in this script, included for future reference.
        Ward_code := LineArray[3]
        
        ;This makes sure the right window is open:
        IfWinNotActive,  Patient Entry
        {
            MsgBox, Please open the  Patient Entry Window
            Return
        }
        
        ;MsgBox, Optional pause: Now doing my thing for for %Episode_out%
        
        WinActivate,  Patient Entry
        WinWaitActive,  Patient Entry
        send, %Episode_out%
        sleep, 400
        send, {Tab}
        sleep, 2000
        WinWaitActive, Patient Entry - Fully Entered
        sleep, 300
        send, {CtrlDown}{Delete}{CtrlUp}
        sleep, 200
        MouseClick, Left, 643, 111, 2, 100
        sleep, 100
        send, %Location_code%
        sleep, 2500
        send, %Ward_code%
        sleep, 200
        send, {Tab}
        sleep, 800
        ;MouseClick, Left, 883, 109, 1 ;Click on Hospital number field start     ;;;;;;;;;;; Use this portion to remove the hospital folder number
        ;sleep, 200
        ;send, {ShiftDown}{End}{ShiftUp}
        ;sleep, 100
        ;send, {Delete}
        ;sleep, 100
        ;send, {Tab}
        ;sleep, 900
        send, {AltDown}u{AltUp}
        sleep, 500
        send, {Enter}
        sleep, 800
     ;Now repeat the process or continue when at the last line
}
return ; Return in AHK will "kill" the script, thus STOP.

ButtonChange_Ref:
sleep, 500
IfWinNotExist, Patient Entry
        {
            MsgBox, Please open the Patient Entry Window
            Return
        }
WinActivate, Patient Entry

FileName := ref_list.csv
Loop, read, ref_list.csv, ;output_list.txt ; output_list.txt is the file to write to, if necessary
{   ;This part of code reads and parses the CSV / txt file and stores results of the read file in an array.
    LineNumber := A_Index
    Loop, parse, A_LoopReadLine, CSV
    
        EpisodeLine := LineNumber
        FileReadLine, Line_out, ref_list.csv, EpisodeLine
        LineArray := StrSplit(Line_out, ",")
        Episode_out := LineArray[1]
        Ref_code := LineArray[2] ; Not actually used in this script, included for future reference.
        
        ;This makes sure the right window is open:
        IfWinNotActive,  Patient Entry
        {
            MsgBox, Please open the  Patient Entry Window
            Return
        }
        
        ;MsgBox, Optional pause: Now doing my thing for for %Episode_out%
        
        WinActivate,  Patient Entry
        WinWaitActive,  Patient Entry
        sleep, 300
        send, %Episode_out%
        sleep, 400
        send, {Tab}
        sleep, 2000
        WinWaitActive, Patient Entry - Fully Entered
        sleep, 900
        ;send, {CtrlDown}{Delete}{CtrlUp}
        ;sleep, 100
        MouseClick, Left, 993, 135, 1
        sleep, 100
        send, {CtrlDown}{ShiftDown}{Home}{ShiftUp}{CtrlUp}
        sleep, 100
        send, %Ref_code%
        sleep, 100
        send, {Tab}
        sleep, 900
        send, {AltDown}u{AltUp}
        sleep, 500
        send, {Enter}
        sleep, 400
        FileAppend, %Episode_out%`,%Ref_code%`,%A_Now%`n, ref_list.txt
        sleep, 200
     ;Now repeat the process or continue when at the last line
}
return ; Return in AHK will "kill" the script, thus STOP.

ButtonRemove_fol:
sleep, 500
IfWinNotExist, Patient Entry
        {
            MsgBox, Please open the Patient Entry Window
            Return
        }
WinActivate, Patient Entry

FileName := folder_list.csv
Loop, read, folder_list.csv, ;output_list.txt ; output_list.txt is the file to write to, if necessary
{   ;This part of code reads and parses the CSV / txt file and stores results of the read file in an array.
    LineNumber := A_Index
    Loop, parse, A_LoopReadLine, CSV
    
        EpisodeLine := LineNumber
        FileReadLine, Line_out, folder_list.csv, EpisodeLine
        LineArray := StrSplit(Line_out, ",")
        Episode_out := LineArray[1]
        Ref_code := LineArray[2] ; Not actually used in this script, included for future reference.
        
        ;This makes sure the right window is open:
        IfWinNotActive,  Patient Entry
        {
            MsgBox, Please open the  Patient Entry Window
            Return
        }
        
        ;MsgBox, Optional pause: Now doing my thing for for %Episode_out%
        
        WinActivate,  Patient Entry
        WinWaitActive,  Patient Entry
        sleep, 300
        send, %Episode_out%
        sleep, 400
        send, {Tab}
        sleep, 2000
        WinWaitActive, Patient Entry - Fully Entered
        sleep, 900
        ;send, {CtrlDown}{Delete}{CtrlUp}
        ;sleep, 100
        MouseClick, Left, 883, 109, 1 ;Click on Hospital number field start
        sleep, 200
        send, {ShiftDown}{End}{ShiftUp}
        sleep, 100
        send, {Delete}
        sleep, 100
        send, {Tab}
        sleep, 900
        send, {AltDown}u{AltUp}
        sleep, 500
        send, {Enter}
        sleep, 400
        FileAppend, %Episode_out%`,%Ref_code%`,%A_Now%`n, folder_list_removed.txt
        sleep, 200
     ;Now repeat the process or continue when at the last line
}
return ; Return in AHK will "kill" the script, thus STOP.

ButtonRe_bill:
sleep, 500
IfWinNotExist, Accounts Re-Pricing Episode
        {
            MsgBox, Please open the Accounts Re-Pricing Window
            Return
        }
WinActivate, Accounts Re-Pricing

FileName := rebill_list.csv
Loop, read, rebill_list.csv, ;output_list.txt ; output_list.txt is the file to write to, if necessary
{
    LineNumber := A_Index
    Loop, parse, A_LoopReadLine, CSV
    
        EpisodeLine := LineNumber
        FileReadLine, Line_out, rebill_list.csv, EpisodeLine
        LineArray := StrSplit(Line_out, ",")
        Episode_out := LineArray[1]
        Test_item := LineArray[2]
        
        IfWinNotActive,  Accounts Re-Pricing Episode
        {
            MsgBox, Please open the  Accounts Re-Pricing Window
            Return
        }
                    
        WinActivate,  Accounts Re-Pricing
        WinWaitActive,  Accounts Re-Pricing
        sleep, 500
        send, %Episode_out%
        sleep, 500
        send, {Tab}
        sleep, 700
        send, {AltDown}a{AltUp}
        sleep, 500
        send, %Episode_out%
        send, {Tab}
        sleep, 700
        send, {AltDown}r{AltUp}
        sleep, 200
        WinWaitActive, Re Billing
        sleep, 400   ; Decrease this time to speed up!!
        send  {AltDown}y{AltUp}
        sleep, 800    ; Decrease this time to speed things up further.
     
}
return

ButtonClear_List:
sleep, 500
IfWinNotExist, Result Entry
        {
            MsgBox, Please open the Result Entry Window
            Return
        }
WinActivate, Result Entry
;^!U::
FileName := clear_list.csv
Loop, read, clear_list.csv, ;output_list.txt ; output_list.txt is the file to write to, if necessary
{
    LineNumber := A_Index
    Loop, parse, A_LoopReadLine, CSV
    
        EpisodeLine := LineNumber
        FileReadLine, Line_out, clear_list.csv, EpisodeLine
        LineArray := StrSplit(Line_out, ",")
        Episode_out := LineArray[1]
        Test_set := LineArray[2]
        
        IfWinActive, Result Entry - Single
            WinClose, Result Entry - Single
        IfWinNotActive, Result Entry
        {
            MsgBox, Please open the Result Entry Window
            Return
        }
        
        ;MsgBox, Now going to clear results for %Episode_out%
        
        WinActivate, Result Entry
        WinWaitActive, Result Entry
        ;Run, Notepad
        ;WinWaitActive, Notepad
        sleep, 500
        send, {AltDown}l{AltUp}
        sleep, 300
        ;MouseClick, Left, 105, 112, 2, 100
        ;MouseClick, Left, 105, 500, 2, 100
        sleep, 1000
        send, %Episode_out%
        sleep, 300
        send, {Tab}
        sleep, 300
        send, {Tab}
        sleep, 300
        send, %Test_set%
        sleep, 200
        send, {Enter}
        sleep, 1000
        MouseClick, Left, 150, 300
        sleep, 1000
        send, {AltDown}e{AltUp}
        WinWaitActive, Result Entry - Single
        
        
        sleep, 3000 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;  <<<<<<<<< Increase this wait time if the connection is slow
        /*
        This portion of code would usually send results
        */
        sleep, 500
        send, {AltDown}l{AltUp}
        WinWaitActive, tkLTResultsEntry (L2016)
        sleep, 300
        send, {Enter}
        sleep, 3000
        WinWaitActive, Result Entry - Single
        sleep, 800
        send, {AltDown}c{AltUp} ; This is to exit Result Entry Window.
        sleep, 1000
        send, {AltDown}l{AltUp}
        sleep, 800
     
}
return

ButtonRe-add_tests:
sleep, 500
IfWinNotExist, Patient Entry
        {
            MsgBox, Please open the Patient Entry Window
            Return
        }
WinActivate, Patient Entry

FileName := re_add_list.csv
Loop, read, re_add_list.csv, ;output_list.txt ; output_list.txt is the file to write to, if necessary
{   ;This part of code reads and parses the CSV / txt file and stores results of the read file in an array.
    LineNumber := A_Index
    Loop, parse, A_LoopReadLine, CSV
    
        EpisodeLine := LineNumber
        FileReadLine, Line_out, re_add_list.csv, EpisodeLine
        LineArray := StrSplit(Line_out, ",")
        Episode_out := LineArray[1]
        Test_set := LineArray[2] 
        
        ;This makes sure the correct window is open:
        IfWinNotActive,  Patient Entry
        {
            MsgBox, Please open the  Patient Entry Window
            Return
        }
        
        ;MsgBox, Optional pause: Now doing my thing for for %Episode_out%
        
        WinActivate,  Patient Entry
        WinWaitActive,  Patient Entry
        send, %Episode_out%
        sleep, 400
        send, {Tab}
        sleep, 2000
        WinWaitActive, Patient Entry - Fully Entered, , 1000
        sleep, 1000
        if ErrorLevel
        {
            MsgBox, Timeout waiting for Patient Entry - Fully Entered Window.`nWhen you press OK the script will resume.
        }
        else
        send, {AltDown}t{AltUp}
        sleep, 200
        MouseClick, Left, 471, 153, 1, 
        sleep, 300
        send, {Delete}
        sleep, 800
        send, {Enter}
        sleep, 500
        ;heleen MouseClick,  Left, 445, 325, 1,
        send, V560
        sleep, 500
        send, {Tab}
        sleep, 500
        send, {Enter}
        sleep, 500
        send, {AltDown}u{AltUp}
        sleep, 500
        send, {Enter}
        sleep, 1200
        FileAppend, %Episode_out%`,%Ref_code%`,%A_Now%`n, re_add_list.txt ; This will append the text file (defined after the last comma in this line) with the log of episodes transcribed.
        
     ;Now repeat the process or continue when at the last line
}
return ; Return in AHK will "kill" the script, thus STOP.

Escape::Reload
;ExitApp
Return

^!r::Reload  ; Assign Ctrl-Alt-R as a hotkey to restart the script.

ButtonClose:
ExitApp

Button_i:
Run, http://github.com/dietervdwes/chemhelp


    

