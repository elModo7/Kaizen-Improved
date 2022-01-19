; KaiZen - Kanji Memorization Game

#NoEnv
#SingleInstance, Force
SetWorkingDir %A_ScriptDir%
SetBatchLines -1
ListLines Off
AutoTrim Off
#KeyHistory 0
voiceId = 2 ; El id de la voz japonés, puede cambiar

Global Mode, Root, Correct, Wrong, Total, Remaining, Errors, Sound, VoiceEnabled, voiceId, RandomNumber, MinIndex, MaxIndex, Kanji, Meaning, CorrectAnswer

IniRead Mode,     KaiZen.ini, Settings, Mode,     1
IniRead Total,    KaiZen.ini, Settings, Total,    25
IniRead MinIndex, KaiZen.ini, Settings, MinIndex, 1
IniRead MaxIndex, KaiZen.ini, Settings, MaxIndex, 50
IniRead FontName, KaiZen.ini, Settings, FontName, Segoe UI
IniRead Sound,    KaiZen.ini, Settings, Sound,    1
IniRead VoiceEnabled, KaiZen.ini, Settings, VoiceEnabled,    1

Menu Tray, Icon, %A_ScriptDir%\Icon\KaiZen.ico

XML := LoadXML(A_ScriptDir . "\Database\Kanji.xml")
Root := XML.selectSingleNode("xml")
Count := Root.ChildNodes.length

If (!Count) {
    MsgBox 0x10, KaiZen, Database Error.
}

Gui KaiZen: New, LabelGui
Gui Color, 0xFEFEFE

Menu GameMenu, Add, New Session`tF2, Restart
Menu GameMenu, Add
Menu GameMenu, Add, Settings`tF3, ShowSettings
Menu GameMenu, Add, Play Sounds, ToggleSound
Menu GameMenu, Add, Enable Voice, ToggleVoice
If (Sound) {
    Menu GameMenu, Check, Play Sounds
}
If (VoiceEnabled) {
    Menu GameMenu, Check, Enable Voice
}
Menu GameMenu, Add
Menu GameMenu, Add, Exit`tEsc, GuiClose

Menu HelpMenu, Add, List of 1000 Kanji, OpenEbook
Menu HelpMenu, Add
Menu HelpMenu, Add, About, ShowAbout

Menu MenuBar, Add, Game, :GameMenu
Menu MenuBar, Add, Help, :HelpMenu
Gui Menu, MenuBar

Gui Add, Tab2, vTab x0 y0 w0 h0 AltSubmit, 1|2|3|4
GuiControl Choose, Tab, %Mode%

WinVer := DllCall("kernel32.dll\GetVersion")
Major  := WinVer & 0xFF
Minor  := WinVer >> 8 & 0xFF
Win8OrLater := Major > 9 || (Major > 5 && Minor > 1)

ButtonStyle := Win8OrLater ? "" : "+0x8000 -Theme" ; BS_FLAT

; Kanji to Meaning mode
Gui Tab, 1
    Gui Font, s84 c0x0099CC Q4, %FontName%
    Gui Add, Text, vKanji x167 y19 w150 h150 +0x201
    Gui Font
    Gui Font, s13 Bold, Segoe UI
    Gui Add, Button, vAlt1 gCheckAnswer x108 y175 w270 h42 %ButtonStyle%
    Gui Add, Button, vAlt2 gCheckAnswer x108 y222 w270 h42 %ButtonStyle%
    Gui Add, Button, vAlt3 gCheckAnswer x108 y269 w270 h42 %ButtonStyle%
    Gui Add, Button, vAlt4 gCheckAnswer x108 y316 w270 h42 %ButtonStyle%
    Gui Add, Button, vAlt5 gCheckAnswer x108 y364 w270 h42 %ButtonStyle%
    Gui Font, s13 Bold, Webdings
    Gui Add, Button, gopenJisho x25 y394 w42 h42, i
    Gui Add, Button, grepeatLastKanji x420 y394 w42 h42, W
    Gui Font

ButtonStyle := Win8OrLater ? "" : "-Theme"

; Meaning to Kanji mode
Gui Tab, 2
    Gui Font, s30 c0x0099CC Bold Q4, Segoe UI
    Gui Add, Text, vMeaning x1 y10 w483 h80 +0x201
    Gui Font
    Gui Font, s42 Q4, %FontName%
    Gui Add, Button, vBtn7 gCheckAnswer x90 y100 w100 h100 %ButtonStyle%
    Gui Add, Button, vBtn8 gCheckAnswer x200 y100 w100 h100 %ButtonStyle%
    Gui Add, Button, vBtn9 gCheckAnswer x310 y100 w100 h100 %ButtonStyle%
    Gui Add, Button, vBtn4 gCheckAnswer x90 y210 w100 h100 %ButtonStyle%
    Gui Add, Button, vBtn5 gCheckAnswer x200 y210 w100 h100 %ButtonStyle%
    Gui Add, Button, vBtn6 gCheckAnswer x310 y210 w100 h100 %ButtonStyle%
    Gui Add, Button, vBtn1 gCheckAnswer x90 y320 w100 h100 %ButtonStyle%
    Gui Add, Button, vBtn2 gCheckAnswer x200 y320 w100 h100 %ButtonStyle%
    Gui Add, Button, vBtn3 gCheckAnswer x310 y320 w100 h100 %ButtonStyle%
    Gui Font

ButtonStyle := Major < 6 ? "" : "+0xE" ; BS_COMMANDLINK

; End of Session
Gui Tab, 3
    Gui Font, s12 cNavy, Segoe UI
    Gui Add, Text, x92 y13 w120 h23 +0x200, End of Session
    Gui Font
    Gui FOnt, s9, Segoe UI
    Gui Add, Text, x92 y45 w120 h23 +0x200, Errors:
    Gui Font, s16, Segoe UI
    Gui Add, ListBox, vErrorList x92 y69 w300 h190
    Gui Font

    If (XP) {
        Gui Font, s12
    }

    Gui Add, Custom, x157 y265 w170 h42 ClassButton %ButtonStyle% gButtonHandler, New Session
    Gui Add, Custom, x157 y309 w170 h42 ClassButton %ButtonStyle% gButtonHandler, Settings
    Gui Add, Custom, x157 y353 w170 h42 ClassButton %ButtonStyle% gButtonHandler, Copy Error List
    Gui Add, Custom, x157 y397 w170 h42 ClassButton %ButtonStyle% gButtonHandler, Exit

; Settings
Gui Tab, 4
    Gui Font, s9, Segoe UI
    Gui Add, GroupBox, x101 y15 w282 h244, Settings
    Gui Add, Text, x120 y35 w120 h23 +0x200, Mode:
    Gui Add, DropDownList, vGameMode x243 y35 w125 AltSubmit, Kanji to Meaning|Meaning to Kanji
    GuiControl Choose, GameMode, %Mode%
    Gui Add, Text, x120 y66 w120 h23 +0x200, Total per session:
    Gui Add, Edit, vTotal x243 y67 w125 Number, %Total%
    Gui Add, GroupBox, x101 y101 w282 h116, Database Range
    Gui Add, Text, x120 y123 w120 h23 +0x200, Minimum index:
    Gui Add, Edit, vMinIndex x243 y124 w125 h21 Number, %MinIndex%
    Gui Add, Text, x120 y151 w120 h23 +0x200, Maximum index:
    Gui Add, Edit, vMaxIndex x243 y152 w125 h21 Number, %MaxIndex%
    Gui Add, Text, x120 y181 w120 h24 +0x200, Kanji in database:
    Gui Add, Edit, x243 y181 w125 h21 Disabled, %Count%
    Gui Add, Text, x120 y225 w120 h23 +0x200, Kanji font:
    Gui Add, ComboBox, vFontName x243 y225 w125, Segoe UI|MS Mincho
    GuiControl Text, FontName, %FontName%

    If (XP) {
        Gui Font, s12
    }

    Gui Add, Custom, x157 y265 w170 h42 ClassButton %ButtonStyle% gButtonHandler, Apply Settings
    Gui Add, Custom, x157 y309 w170 h42 ClassButton %ButtonStyle% gButtonHandler, Apply and Restart
    Gui Add, Custom, x157 y353 w170 h42 ClassButton %ButtonStyle% gButtonHandler, Reset Settings
    Gui Add, Custom, x157 y397 w170 h42 ClassButton %ButtonStyle% gButtonHandler, Return

Gui Font, s9, Segoe UI
Gui Add, Statusbar
SB_SetParts(120, 120, 120, 125)

Gui Show, w485 h477, KaiZen - 漢字検定

OnMessage(0x100, "OnWM_KEYDOWN")
OnMessage(0x6,   "OnWM_ACTIVATE")
OnMessage(0x135, "OnWM_CTLCOLORBTN")

Start:
    Correct := 0
    Wrong := 0
    Remaining := Total
    Errors := ""
    
    ResetStatusBar()

    Global Id := [], AvailableId := []

    Loop % (MaxIndex - MinIndex + 1) {
        Index := MinIndex + A_Index - 1
        Id[A_Index] := Index
        AvailableId[A_Index] := Index
    }

    If (Mode == 1) {
        GoSub KanjiToMeaning
    } Else {
        GoSub MeaningToKanji
    }
Return

sayKanji(kanjiToSpeak){
    global lastKanji
    lastKanji := kanjiToSpeak
    sayLastKanji()
}

sayLastKanji(async := 1){
    global voiceId, v, lastKanji
    v := ComObjCreate("SAPI.SpVoice")
    v.Voice := v.GetVoices().Item(voiceId) ; Item is Zero based
    v.rate := -1 ; slow down speak
    v.Speak(lastKanji, async) ; announce String
}

MeaningToKanji:
    GuiControl Focus, msctls_statusbar321

    Random RandomNumber, 1, AvailableId.Length()

    ; Get the meaning and the correct answer
    Meaning := GetMeaning(AvailableId[RandomNumber])
    GuiControl, KaiZen:, Meaning, %Meaning%
    CorrectAnswer := GetKanji(AvailableId[RandomNumber])
    if(VoiceEnabled){
        sayKanji(CorrectAnswer)
    }

    AvailableId.Remove(RandomNumber)

    Alternatives := []
    Counter := 0
    ; Get the 9 alternatives
    Loop {
        Random RandomNumber, 1, Id.Length()

        Kanji := GetKanji(Id[RandomNumber])

        If (Kanji == CorrectAnswer) {
            Continue
        }

        ; Prevent repeated alternatives
        Repeated := False
        Loop % Alternatives.Length() {
            If (Alternatives[A_Index] == Kanji) {
                Repeated := True
                Break
            }
        }
        If (Repeated) {
            Continue
        }

        Counter++
        Alternatives[Counter] := Kanji

        If (Counter == 9) {
            Break
        }
    }
    
    ; Rewrite one of the alternatives with the correct answer
    Random RandomNumber, 1, 9
    Alternatives[RandomNumber] := CorrectAnswer
    
    ; Fill the buttons with alternatives
    Loop 9 {
        GuiControl, KaiZen:, Btn%A_Index%, % Alternatives[A_Index]
    }
Return

KanjiToMeaning:
    GuiControl Focus, msctls_statusbar321

    Random RandomNumber, 1, AvailableId.Length()

    Kanji := GetKanji(AvailableId[RandomNumber])
    GuiControl, KaiZen:, Kanji, %Kanji%
    if(VoiceEnabled){
        sayKanji(Kanji)
    }
    CorrectAnswer := GetMeaning(AvailableId[RandomNumber])
    StringUpper CorrectAnswer, CorrectAnswer

    AvailableId.Remove(RandomNumber)

    Alternatives := []
    Counter := 0
    ; Get the 5 alternatives
    Loop {
        Random RandomNumber, 1, Id.Length()

        Meaning := GetMeaning(Id[RandomNumber])
        StringUpper Meaning, Meaning

        If (Meaning == CorrectAnswer) {
            Continue
        }

        ; Prevent repeated alternatives
        Repeated := False
        Loop % Alternatives.Length() {
            If (Alternatives[A_Index] == Meaning) {
                Repeated := True
                Break
            }
        }
        If (Repeated) {
            Continue
        }

        Counter++
        Alternatives[Counter] := Meaning

        If (Counter == 5) {
            Break
        }
    }
    
    ; Rewrite one of the alternatives with the correct answer
    Random RandomNumber, 1, 5
    Alternatives[RandomNumber] := CorrectAnswer
    
    ; Fill the buttons with alternatives
    Loop 5 {
        GuiControl, KaiZen:, Alt%A_Index%, % Alternatives[A_Index]
    }
Return

CheckAnswer:
    GuiControl -Default, %A_GuiControl%
    GuiControlGet Answer,, %A_GuiControl%
    CheckAnswer(Answer)
Return

CheckAnswer(Answer) {
    Menu GameMenu, Disable, New Session`tF2

    If (Answer == CorrectAnswer
    || (Mode == 2 && Root.selectSingleNode("//kanji[@key=""" . Answer . """]").FirstChild.text == Meaning)) {
        if(VoiceEnabled){
            sayLastKanji(0)
        }else{
            PlaySound("Correct")
        }
        Correct++
        SB_SetText("Correct: " . Correct, 1)
        UpdatePercent()
    } Else {
        PlaySound("Wrong")
        BlinkBorder(Answer, "Red", 1000)
        Sleep 200
        Wrong++
        SB_SetText("Wrong: " . Wrong, 2)
        UpdatePercent()

        If (Mode == 1) {
            StringLower CorrectAnswer, CorrectAnswer
            Errors .= Kanji . " - " . CorrectAnswer . "|"
        } Else {
            Errors .= CorrectAnswer . " - " . Meaning . "|"
        }
        
        ButtonPrefix := (Mode == 1) ? "Alt" : "Btn"
        BlinkBorder(ButtonPrefix . RandomNumber, 0x07CAEA, 3000)
    }

    Remaining--
    SB_SetText("Remaining: " . Remaining, 4)

    Menu GameMenu, Enable, New Session`tF2

    If (!Remaining) {
        PlaySound("End")
        ShowErrors()
        Return
    }

    If (Mode == 1) {
        GoSub KanjiToMeaning
    } Else {
        GoSub MeaningToKanji
    }
}

UpdatePercent() {
    Percent := (Correct / (Correct + Wrong)) * 100
    SB_SetText("Percent: " . Round(Percent) . "%", 3)
}

OnWM_KEYDOWN(wParam, lParam, msg, hWnd) {
    Local Answer

    GuiControlGet Tab,, Tab
    If (Mode == 1 && Tab == 1) {
        Char := Chr(wParam > 96 ? wParam - 48 : wParam)
        If Char in 1,2,3,4,5
        {
            GuiControlGet Answer,, Alt%Char%
            CheckAnswer(Answer)
            Return False
        }
    } Else If (Mode == 2 && Tab == 2) {
        ; Numpad keys 1-9 (97-105)
        If (wParam >= 97 && wParam <= 105) {
            ButtonIndex := wParam - 96
            GuiControlGet Answer,, Btn%ButtonIndex%
            CheckAnswer(Answer)
            Return False
        }
    }

    GuiControlGet vVar, KaiZen: Name, %hWnd%
    If (vVar == "Tab") {
        Return False
    }
}

OnWM_ACTIVATE() {
    GuiControl Focus, msctls_statusbar321
}

BlinkBorder(ClassNN, Color, Duration, r := 5) {
    Local X, Y, W, H, Index
    
    WinGetPos wX, wY
    GuiControlGet hWnd, hWnd, %ClassNN%
    ControlGetPos X, Y, W, H,, ahk_id %hWnd%
    X += wX
    Y += wY

    Loop 4 {
        Index := A_Index + 80
        Gui, %Index%: -Caption ToolWindow AlwaysOnTop
        Gui, %Index%: Color, %Color%
    }

    Gui, 81: Show, % "NA X" (X - r) " Y" (Y - r) " W" (W + r + r) " H" r
    Gui, 82: Show, % "NA X" (X - r) " Y" (Y + H) " W" (W + r + r) " H" r
    Gui, 83: Show, % "NA X" (X - r) " Y" Y " W" r " H" H
    Gui, 84: Show, % "NA X" (X + W) " Y" Y " W" r " H" H

    Sleep %Duration%

    Loop 4 {
        Index := A_Index + 80
        Gui, %Index%: Destroy
    }
}

ResetStatusBar() {
    Gui KaiZen: Default
    SB_SetText("Correct: 0", 1)
    SB_SetText("Wrong: 0", 2)
    SB_SetText("Percent: 0%", 3)
    SB_SetText("Remaining: " . Total, 4)
}

ShowErrors() {
    GuiControl,, ErrorList, |%Errors%
    GuiControl Choose, Tab, 3
}

CopyErrors:
    Temp := ""
    Loop Parse, Errors, |
    {
        Temp .= A_LoopField . "`r`n"
    }

    Clipboard := RTrim(Temp, "`r`n")

    MsgBox 0x40, KaiZen, Error list copied to the clipboard.
Return

Restart:
    GuiControlGet Mode,, GameMode
    GuiControl Choose, Tab, %Mode%
    GoSub Start
Return

GuiEscape:
GuiClose:
    IniWrite %Sound%, KaiZen.ini, Settings, Sound
    IniWrite %VoiceEnabled%, KaiZen.ini, Settings, VoiceEnabled
    ExitApp

ShowSettings:
    GuiControl Choose, Tab, 4
Return

SaveSettings:
SaveAndRestart:
    Gui KaiZen: Submit, NoHide

    If (MinIndex == 0) {
        MsgBox 0x10, KaiZen, Minimum index must be higher than zero.
        Return
    }

    If (MaxIndex < MinIndex || (MaxIndex - MinIndex< 10)) {
        MsgBox 0x10, KaiZen, Invalid database range.
        Return
    }

    If (MaxIndex - MinIndex + 1 < Total) {
        MsgBox 0x10, KaiZen, The number of kanji per session exceeds the database range specified.
        Return
    }

    If (MaxIndex > Root.ChildNodes.length) {
        MsgBox 0x10, KaiZen, The maximum index exceeds the database length.
        Return
    }

    Font := StrReplace(FontName, " (Bold)", "", Bold, 1)
    FontStyle := Bold ? "Bold" : ""
    Gui Font
    If (Mode == 1) {
        Gui Font, s84 c0x0099CC %FontStyle% Q4, %Font%
        GuiControl Font, Kanji
    } Else {
        Gui Font, s42 %FontStyle% Q4, %Font%
        Loop 9 {
            GuiControl Font, Btn%A_Index%
        }
    }

    IniWrite %GameMode%, KaiZen.ini, Settings, Mode
    IniWrite %Total%,    KaiZen.ini, Settings, Total
    IniWrite %MinIndex%, KaiZen.ini, Settings, MinIndex
    IniWrite %MaxIndex%, KaiZen.ini, Settings, MaxIndex
    IniWrite %FontName%, KaiZen.ini, Settings, FontName
    IniWrite %VoiceEnabled%, KaiZen.ini, Settings, VoiceEnabled

    If (!Remaining || A_ThisLabel == "SaveAndRestart") {
        GoSub Restart
    } Else {
        GuiControl Choose, Tab, %Mode%
    }
Return

ResetSettings:
    GuiControl,, Total, 25
    GuiControl,, MinIndex, 1
    GuiControl,, MaxIndex, 50
    GuiControl Choose, Font, Segoe UI
    GuiControl,, Sound, 1
Return

GuiContextMenu:
    GuiControlGet CtrlText,, %A_GuiControl%
    Menu ContextMenu, Add, Copy, Copy
    Menu ContextMenu, Show, %A_GuiX%, %A_GuiY%
Return

Copy:
    Clipboard := CtrlText
Return

LoadXML(FileName) {
    Local x
    x := ComObjCreate("MSXML2.DOMDocument.6.0")
    x.async := False
    x.load(FileName)
    Return x
}

GetKanji(Index) {
    Return Root.childNodes[Index - 1].getAttribute("key")
}

GetMeaning(Index) {
    Return Root.childNodes[Index - 1].FirstChild.text
}

PlaySound(SoundType) {
    If (Sound) {
        SoundPlay %A_ScriptDir%\Sounds\%SoundType%.wav
    }
}

ToggleSound:
    Sound := !Sound
    Menu GameMenu, ToggleCheck, Play Sounds
Return

ToggleVoice:
    VoiceEnabled := !VoiceEnabled
    Menu GameMenu, ToggleCheck, Enable Voice
    if(VoiceEnabled){
        sayKanji(CorrectAnswer)
    }
Return

OpenEbook:
    Run %A_ScriptDir%\Help\List of 1000 Kanji.pdf
Return

ShowAbout:
    OnMessage(0x44, "OnMsgBox")
    Gui +OwnDialogs
    MsgBox 0x80, About, KaiZen v1.0.0`nKanji Memorization Game
    OnMessage(0x44, "")
Return

OnMsgBox() {
    DetectHiddenWindows On
    Process Exist
    If (WinExist("ahk_class #32770 ahk_pid " . ErrorLevel)) {
        hIcon := LoadPicture(A_ScriptDir . "\Icon\KaiZen.ico", "w32 Icon1", _)
        SendMessage 0x172, 1, %hIcon% , Static1 ; STM_SETIMAGE
    }
}

OnWM_CTLCOLORBTN() {
    Static Brush := DllCall("Gdi32.dll\CreateSolidBrush", "UInt", 0xFFFFFF, "UPtr")
    Return Brush
}

ButtonHandler:
    If (A_GuiEvent == "Normal") {
        If (A_GuiControl == "New Session") {
            GoSub Restart
        } Else If (A_GuiControl == "Settings") {
            GoSub ShowSettings
        } Else If (A_GuiControl == "Copy Error List") {
            GoSub CopyErrors
        } Else If (A_GuiControl == "Exit") {
            GoSub GuiClose
        } Else If (A_GuiControl == "Apply Settings") {
            GoSub SaveSettings
        } Else If (A_GuiControl == "Apply and Restart") {
            GoSub SaveAndRestart
        } Else If (A_GuiControl == "Reset Settings") {
            GoSub ResetSettings
        } Else If (A_GuiControl == "Return") {
            If (!Remaining) {
                GuiControl Choose, Tab, 3 ; Return to End of Session
            } Else {
                GuiControl Choose, Tab, %Mode%
            }
        }
    }
Return

openJisho:
    ;~ Run, % A_ScriptDir "/lib/web_embebida.ahk https://jisho.org/search/" Kanji ; Deprecated
    if(!pipa)
        Gosub,init
    url = % "https://jisho.org/search/" Kanji
    WB.Navigate(url)
    loop
      If !WB.busy
         break
    gui,WB:show
return

repeatLastKanji:
    sayLastKanji()
return

init:
    Gui, WB:+LastFound +Resize +OwnDialogs
    Gui, WB:Add, ActiveX, w1280 h600 x0 y0 vWB hwndATLWinHWND, Shell.Explorer
    WB.silent := true
    IOleInPlaceActiveObject_Interface:="{00000117-0000-0000-C000-000000000046}"
    pipa := ComObjQuery(WB, IOleInPlaceActiveObject_Interface)
    OnMessage(WM_KEYDOWN:=0x0100, "WM_KEYDOWN")
    OnMessage(WM_KEYUP:=0x0101, "WM_KEYDOWN")
    gui,WB:show, w1280 h600 , Jisho Search
return

WBGuiSize:
    WinMove, % "ahk_id " . ATLWinHWND, , 0,0, A_GuiWidth, A_GuiHeight
return

WBGuiClose:
terminate:
    Gui, WB:Hide
return



WM_KEYDOWN(wParam, lParam, nMsg, hWnd)
{
   global pipa
   static keys:={9:"tab", 13:"enter", 46:"delete", 38:"up", 40:"down"}
   if keys.HasKey(wParam)
   {
      WinGetClass, ClassName, ahk_id %hWnd%
      if  (ClassName = "Internet Explorer_Server")
      {
      ;// Build MSG Structure
         VarSetCapacity(Msg, 48)
         for i,val in [hWnd, nMsg, wParam, lParam, A_EventInfo, A_GuiX, A_GuiY]
            NumPut(val, Msg, (i-1)*A_PtrSize)
      ;// Call Translate Accelerator Method
         TranslateAccelerator := NumGet(NumGet(1*pipa)+5*A_PtrSize)
         DllCall(TranslateAccelerator, "Ptr",pipa, "Ptr",&Msg)
         return, 0
      }
   }
}