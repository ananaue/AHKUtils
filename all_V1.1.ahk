#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;MISC

; Open OneDrive with RAlt + o
>!o:: Run, C:\Users\USER\OneDrive - University of Example

; RAlt + [, or .] for Back and Forward buttons 
; (For ex, when on a laptop and using file explorer)
>!,::Browser_Back
>!.::Browser_Forward

; These change Office sub and supercript hotkeys to RCtrl + \ or Backspace respectively (I think the new Word hotkeys are ok, but the old ones..)
; >^\::
;     if WinActive("ahk_exe WINWORD.EXE")
;     {
;         Send, ^+-
;         return
;     }
;     else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
;     {
;         Send, ^=
;     }
; return

; >^BS::
;     if WinActive("ahk_exe WINWORD.EXE")
;     {
;         Send, ^+=
;         return
;     }
;     else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
;     {
;         Send, ^+=
;     }
; return

; Shift + Mousewheel for horizontal scroll in some programs (Uncomment lines to make applicable to only Onenote)
; Excel is still Shift + Ctrl + Mousewheel
;if WinActive("ahk_exe ONENOTE.EXE") 
;{
+WheelDown::WheelRight
+WheelUp::WheelLeft
;}
;return

; Middle Click folder to open in new window
; Written by lexikos on AHK forum - https://www.autohotkey.com/boards/viewtopic.php?t=9832
MouseIsOver(WinTitle) {
    MouseGetPos,,, Win
    return WinExist(WinTitle . " ahk_id " . Win)
}
#If MouseIsOver("ahk_class CabinetWClass")
MButton::
    Send {LButton}{AppsKey}e
    return
#If

; In Word, switch to one-page view (for when zoomed out) with RAlt+l
#if WinActive("ahk_exe WINWORD.EXE")
>!l::
    Send, !w
    Send, 1
    Send, !h
    Send, {Enter}
return
#if

; Toggle current window always-on-top
; Adapted from answer by Ethan Malloy - https://stackoverflow.com/a/75881915
>^Enter::
    WinSet, AlwaysOnTop , -1, A
return

;MATH

>^v::  ; square root
    Send, √
return

>^+v::  ; check mark
    Send, ✓
return

>^`::  ; degree
    Send, °
return

>^+`::  ; bullet
    Send, •
return

>^'::  ; another bullet
    Send, •
return

>+NumpadMult:: ; another degree (RShift + numpad asterisk)
    Send, °
return

>^+BS::  ; approximately equal
    Send, ≈
return

>^+\::  ; plus or minus
    Send, ±
return

>^!2::  ; half
    Send, ½
return

; For squared and cubed, a unicode ²/³ will be used except in Word, Onenote, Powerpoint, or Outlook.
; There, a superscripted 2 will be inserted.
; Similar for superscript 3, -, +, etc
; and subscript 0, 1, 2, 3, etc

>^!0::
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^+=
        Send, 0
        Send, ^+=
        return
    }
    else
    {
        Send, ⁰
    }
return

>^1::
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^+=
        Send, 1
        Send, ^+=
        return
    }
    else
    {
        Send, ¹
    }
return

>^2::
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^+=
        Send, 2
        Send, ^+=
        return
    }
    else
    {
        Send, ²
    }
return

>^3::
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^+=
        Send, 3
        Send, ^+=
        return
    }
    else
    {
        Send, ³
    }
return

>^!4::
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^+=
        Send, 4
        Send, ^+=
        return
    }
    else
    {
        Send, ⁴
    }
return

>^!5::
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^+=
        Send, 5
        Send, ^+=
        return
    }
    else
    {
        Send, ⁵
    }
return

>^!6::
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^+=
        Send, 6
        Send, ^+=
        return
    }
    else
    {
        Send, ⁶
    }
return

>^!7::
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^+=
        Send, 7
        Send, ^+=
        return
    }
    else
    {
        Send, ⁷
    }
return

>^!8::
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^+=
        Send, 8
        Send, ^+=
        return
    }
    else
    {
        Send, ⁸
    }
return

>^!9::
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^+=
        Send, 9
        Send, ^+=
        return
    }
    else
    {
        Send, ⁹
    }
return

>^!n::
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^+=
        Send, n
        Send, ^+=
        return
    }
    else
    {
        Send, ⁿ
    }
return

; only LShift (capital theta is RShift)
>^<+0::
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send, ^+-
        Send, 0
        Send, ^+-
        return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^=
        Send, 0
        Send, ^=
        return
    }
    else
    {
        Send, ₀
    }
return

>^+1::
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send, ^+-
        Send, 1
        Send, ^+-
        return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^=
        Send, 1
        Send, ^=
        return
    }
    else
    {
        Send, ₁
    }
return

>^+2::
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send, ^+-
        Send, 2
        Send, ^+-
        return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^=
        Send, 2
        Send, ^=
        return
    }
    else
    {
        Send, ₂
    }
return

>^+3::
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send, ^+-
        Send, 3
        Send, ^+-
        return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^=
        Send, 3
        Send, ^=
        return
    }
    else
    {
        Send, ₃
    }
return

; only LShift (capital psi is RShift)
>^<+4::
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send, ^+-
        Send, 4
        Send, ^+-
        return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^=
        Send, 4
        Send, ^=
        return
    }
    else
    {
        Send, ₄
    }
return

>^+5::
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send, ^+-
        Send, 5
        Send, ^+-
        return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^=
        Send, 5
        Send, ^=
        return
    }
    else
    {
        Send, ₅
    }
return

>^+6::
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send, ^+-
        Send, 6
        Send, ^+-
        return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^=
        Send, 6
        Send, ^=
        return
    }
    else
    {
        Send, ₆
    }
return

>^+7::
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send, ^+-
        Send, 7
        Send, ^+-
        return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^=
        Send, 7
        Send, ^=
        return
    }
    else
    {
        Send, ₇
    }
return

>^+8::
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send, ^+-
        Send, 8
        Send, ^+-
        return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^=
        Send, 8
        Send, ^=
        return
    }
    else
    {
        Send, ₈
    }
return

>^+9::
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send, ^+-
        Send, 9
        Send, ^+-
        return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^=
        Send, 9
        Send, ^=
        return
    }
    else
    {
        Send, ₉
    }
return

>^+n::
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send, ^+-
        Send, n
        Send, ^+-
        return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^=
        Send, n
        Send, ^=
        return
    }
    else
    {
        Send, ₙ
    }
return

>^-::
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^+=
        Send, -
        Send, ^+=
        return
    }
    else
    {
        Send, ⁻
    }
return

>^=::
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^+=
        Send, +=
        Send, ^+=
        return
    }
    else
    {
        Send, ⁺
    }
return


>^+i::
if WinActive("ahk_exe WINWORD.EXE")
    {
        Send, ^+-
        Send, i
        Send, ^+-
        return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^=
        Send, i
        Send, ^=
        return
    }
    else
    {
        Send, ᵢ
    }
return

; only LShift (capital xi is RShift)
>^<+x::
if WinActive("ahk_exe WINWORD.EXE")
    {
        Send, ^+-
        Send, x
        Send, ^+-
        return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send, ^=
        Send, x
        Send, ^=
        return
    }
    else
    {
        Send, ₓ
    }
return


; I don't think this is an unicode characters for sub f,
; so this only works in Word/Onenote/Powerpoint/Outlook.
#if WinActive("ahk_exe WINWORD.EXE")

; only LSHift (capital phi is RShift)
>^<+f::
    Send, ^+-
    Send, f
    Send, ^+-
return
#if

#if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
>^<+f::
    Send, ^=
    Send, f
    Send, ^=
return
#if

>^6::   ; another degree
    Send, °
return

>^8::  ; infinity
    Send, ∞
return

;GREEK
; Lowercase

>^a::  ; alpha
    Send, α
return

>^b::  ; beta
    Send, β
return

>^g::  ; gamma
    Send, γ
return

>^d::  ; delta
    Send, δ
return

>^e::  ; epsilon
    Send, ε
return

>^z::  ; zeta
    Send, ζ
return

>^h::  ; eta
    Send, η
return

>^0::  ; theta  ; unconventional
   Send, θ
return

>^i::  ; iota
   Send, ι
return

>^k::  ; kappa
   Send, κ
return

>^l::  ; lambda
    Send, λ
return

>^m::  ; mu
    Send, μ
return

>^n::  ; nu
    Send, ν
return

>^x::  ; xi
    Send, ξ
return

>^o::  ; omicron
    Send, ο
return

>^p::  ; pi
    Send, π
return

>^r::  ; rho
    Send, ρ
return

>^s::  ; sigma
    Send, σ
return

>^t::  ; tau
    Send, τ
return

>^y::  ; upsilon
    Send, υ
return

>^f::  ; phi
    Send, φ
return

>^c::  ; chi
    Send, χ
return

>^4::  ; psi  ; ($) unconventional
   Send, ψ
return

>^w::  ; omega
    Send, ω
return

; Uppercase

>^+g::  ; Gamma
    Send, Γ
return

>^+d::  ; Delta
    Send, Δ
return

; only RShift
>^>+0::  ; Theta  ; unconventional
   Send, Θ
return

>^+l::  ; Lambda
   Send, Λ
return

; only Rshift
>^>+x::  ; Xi
  Send, Ξ
return

>^+p::  ; Pi
   Send, Π
return

>^+s::  ; Sigma
  Send, Σ
return

; only RShift
>^>+f::  ; Phi
  Send, Φ
return

; only RShift
>^>+4::  ; Psi
  Send, Ψ
return

>^+w::  ; Omega
  Send, Ω
return

