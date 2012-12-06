#NoEnv
#SingleInstance force
keyDelay := 25
i := 0
OutputVar := ""
ItemRecordIterations := 10
controlDelay := 40
directReaderOn := ""
checkDirectReader := ""
StartTime := ""
EndTime := ""
SetupMsgBox := ""
SetkeyDelay, %keyDelay%
SetTitleMatchMode, 1
SetTitleMatchMode, Fast
SetControlDelay, %controlDelay%

WinActivate, Check In - Bulk - Polaris
Sleep, 250
ControlFocus, SysListView323, Check In - Bulk - Polaris
Send {Home}
Send +{End}
;PostMessage, 0x185, 1, -1, SysListView323, Check In - Bulk - Polaris