#NoEnv
#SingleInstance force
keydelay = 25
i = 0
OutputVar = ""
ItemRecordIterations = 5
SetKeyDelay, %keydelay%
SetTitleMatchMode, 1
SetTitleMatchMode, Fast
;SetControlDelay, 40
SetControlDelay -1

;********************************
;
; # Windows Key
; ! Alt
; ^ Control
; + Shift
;
;********************************

;********************************
;
; DirectReader
; Control + Alt + d
;
;********************************

^!+d::
{
	WinActivate, CircControl-DirectReader
	ControlClick, OFF, CircControl-DirectReader
	ExitApp
}