^!z::  ; Control+Alt+Z hotkey.
MouseGetPos, MouseX, MouseY
PixelGetColor, color, %MouseX%, %MouseY%
MsgBox The color at the current cursor position is %MouseX%, %MouseY%.
FormatTime, TimeString
FormatTime, TimeString,, Time
MsgBox The current time is %TimeString%.
FormatTime, TimeString, T12, Time
MsgBox The current 24-hour time is %TimeString%.
return