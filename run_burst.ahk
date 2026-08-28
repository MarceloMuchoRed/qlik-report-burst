#Requires AutoHotkey v2.0
#SingleInstance Force

; Ctrl+G runs the Qlik report burst (via run_burst.bat, next to this script).
; The hotkey is global. Note: Ctrl+G is also "Go To"/"Find Next" in some apps
; (Excel, editors); if that clashes, change the ^g below to another combo.
;
; This .ahk must be running for the hotkey to work — double-click it (AutoHotkey
; v2 must be installed) and it sits in the tray. To start it automatically on
; login, put a shortcut to this file in your Startup folder
; (Win+R -> shell:startup).

^g:: {
    Run('"' A_ScriptDir '\run_burst.bat"')
}
