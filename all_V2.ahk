#Requires AutoHotkey v2.0
; Converted from AHK v1.1 using github.com/mmikeww/AHK-v2-script-converter

;MISC

; Open OneDrive with RAlt + o
>!o::Run("C:\Users\USER\OneDrive - University of Example")

; RAlt + [, or .] for Back and Forward buttons 
; (For ex, when on a laptop and using file explorer)
>!,::Browser_Back
>!.::Browser_Forward

; These change Office sub and supercript hotkeys to RCtrl + \ or Backspace respectively (I think the new hotkeys are ok, but the old ones..)
; >^\::
; {
;     if WinActive("ahk_exe WINWORD.EXE")
;     {
;         Send("^+-")
;     }
;     else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
;     {
;         Send("^=")
;     }
; }

; >^BS::
; {
;     if WinActive("ahk_exe WINWORD.EXE")
;     {
;         Send("^+=")
;     }
;     else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
;     {
;         Send("^+=")
;     }
; }


; Shift + Mousewheel for horizontal scroll in some programs (Uncomment lines to make applicable to only Onenote)
; Excel is still Shift + Ctrl + Mousewheel
;if WinActive("ahk_exe ONENOTE.EXE") 
;{
+WheelDown::WheelRight
+WheelUp::WheelLeft
;}

; Middle Click folder to open in new window
; Written by lexikos on AHK forum - https://www.autohotkey.com/boards/viewtopic.php?t=9832
MouseIsOver(WinTitle) {
    MouseGetPos(, , &Win)
    return WinExist(WinTitle . " ahk_id " . Win)
}
#HotIf MouseIsOver("ahk_class CabinetWClass")
MButton::
{
    Send("{LButton}{AppsKey}e")
}
#HotIf

; In Word, switch to one-page view (for when zoomed out) with RAlt+l
#HotIf WinActive("ahk_exe WINWORD.EXE")
>!l::
{
    Send("!w")
    Send(1)
    Send("!h")
    Send("{Enter}")
}
#HotIf

; Toggle current window always-on-top
; Written by Ethan Malloy on Stack Overflow - https://stackoverflow.com/a/75881915
^>Enter::
{
    WinSetAlwaysOnTop -1, "A"
}

;MATH

>^v::  ; square root
{
    Send("√")
}

>^+v::  ; check mark
{
    Send("✓")
}

>^`::  ; degree
{
    Send("°")
}

>^+`::  ; bullet
{
    Send("•")
}

>^'::  ; another bullet
{
    Send("•")
}

>+NumpadMult:: ; another degree (RShift + numpad asterisk)
{
    Send("°")
}

; The shifts in these two can be removed if not in use for something else.
>^+BS::  ; approximately equal
{
    Send("≈")
}

>^+\::  ; plus or minus
{
    Send("±")
}

>^!2::  ; half
{
    Send("½")
}

; For squared and cubed, a unicode ²/³ will be used except in Word, Onenote, Powerpoint, or Outlook.
; There, a superscripted 2 will be inserted.
; Similar for superscript 3, -, +, etc
; and subscript 0, 1, 2, 3, etc

>^!0::
{
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^+=")
        Send("0")
        Send("^+=")
        Return
    }
    else
    {
        Send("⁰")
    }
}

>^1::
{
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^+=")
        Send("1")
        Send("^+=")
        Return
    }
    else
    {
        Send("¹")
    }
}

>^2::
{
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^+=")
        Send("2")
        Send("^+=")
        Return
    }
    else
    {
        Send("²")
    }
}

>^3::
{
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^+=")
        Send("3")
        Send("^+=")
        Return
    }
    else
    {
        Send("³")
    }
}

>^!4::
{
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^+=")
        Send(4)
        Send("^+=")
        Return
    }
    else
    {
        Send("⁴")
    }
}

>^!5::
{
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^+=")
        Send(5)
        Send("^+=")
        Return
    }
    else
    {
        Send("⁵")
    }
}

>^!6::
{
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^+=")
        Send(6)
        Send("^+=")
        Return
    }
    else
    {
        Send("⁶")
    }
}

>^!7::
{
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^+=")
        Send(7)
        Send("^+=")
        Return
    }
    else
    {
        Send("⁷")
    }
}

>^!8::
{
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^+=")
        Send(8)
        Send("^+=")
        Return
    }
    else
    {
        Send("⁸")
    }
}

>^!9::
{
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^+=")
        Send(9)
        Send("^+=")
        Return
    }
    else
    {
        Send("⁹")
    }
}

>^!n::
{
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^+=")
        Send("n")
        Send("^+=")
        Return
    }
    else
    {
        Send("ⁿ")
    }
}


; only LShift (capital theta is RShift)
>^<+0::
{
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send("^+-")
        Send("0")
        Send("^+-")
        Return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^=")
        Send("0")
        Send("^=")
        Return
    }
    else
    {
        Send("₀")
    }
}

>^+1::
{
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send("^+-")
        Send("1")
        Send("^+-")
        Return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^=")
        Send("1")
        Send("^=")
        Return
    }
    else
    {
        Send("₁")
    }
}

>^+2::
{
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send("^+-")
        Send("2")
        Send("^+-")
        Return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^=")
        Send("2")
        Send("^=")
        Return
    }
    else
    {
        Send("₂")
    }
}

>^+3::
{
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send("^+-")
        Send("3")
        Send("^+-")
        Return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^=")
        Send("3")
        Send("^=")
        Return
    }
    else
    {
        Send("₃")
    }
}

; only LShift (capital psi is RShift)
>^<+4::
{
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send("^+-")
        Send("4")
        Send("^+-")
        Return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^=")
        Send("4")
        Send("^=")
        Return
    }
    else
    {
        Send("₄")
    }
}

>^+5::
{
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send("^+-")
        Send("5")
        Send("^+-")
        Return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^=")
        Send("5")
        Send("^=")
        Return
    }
    else
    {
        Send("₅")
    }
}

>^+6::
{
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send("^+-")
        Send("6")
        Send("^+-")
        Return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^=")
        Send("6")
        Send("^=")
        Return
    }
    else
    {
        Send("₆")
    }
}

>^+7::
{
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send("^+-")
        Send(7)
        Send("^+-")
        Return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^=")
        Send(7)
        Send("^=")
        Return
    }
    else
    {
        Send("₇")
    }
}

>^+8::
{
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send("^+-")
        Send(8)
        Send("^+-")
        Return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^=")
        Send(8)
        Send("^=")
        Return
    }
    else
    {
        Send("₈")
    }
}

>^+9::
{
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send("^+-")
        Send(9)
        Send("^+-")
        Return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^=")
        Send(9)
        Send("^=")
        Return
    }
    else
    {
        Send("₉")
    }
}

>^+n::
{
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send("^+-")
        Send("n")
        Send("^+-")
        Return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^=")
        Send("n")
        Send("^=")
        Return
    }
    else
    {
        Send("ₙ")
    }
}

>^-::
{
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^+=")
        Send("-")
        Send("^+=")
        Return
    }
    else
    {
        Send("⁻")
    }
}

>^=::
{
    if WinActive("ahk_exe WINWORD.EXE") or WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^+=")
        Send("+=")
        Send("^+=")
        Return
    }
    else
    {
        Send("⁺")
    }
}

; I don't think there is an unicode character for sub f,
; so this only works in Word/Onenote/Powerpoint/Outlook.
#HotIf WinActive("ahk_exe WINWORD.EXE")

; only LSHift (capital phi is RShift)
>^<+f::
{
    Send("^+-")
    Send("f")
    Send("^+-")
}
#HotIf

#HotIf WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")

; only LSHift (capital phi is RShift)
>^<+f::
{
    Send("^=")
    Send("f")
    Send("^=")
}
#HotIf

>^+i::
{
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send("^+-")
        Send("i")
        Send("^+-")
        Return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^=")
        Send("i")
        Send("^=")
        Return
    }
    else
    {
        Send("ᵢ")
    }
}

; only LShift (capital xi is RShift)
>^<+x::
{
    if WinActive("ahk_exe WINWORD.EXE")
    {
        Send("^+-")
        Send("x")
        Send("^+-")
        Return
    }
    else if WinActive("ahk_exe ONENOTE.EXE") or WinActive("ahk_exe POWERPNT.EXE") or WinActive("ahk_exe OUTLOOK.EXE")
    {
        Send("^=")
        Send("x")
        Send("^=")
        Return
    }
    else
    {
        Send("ₓ")
    }
}


>^6::   ; another degree
{ 
    Send("°")
} 

>^8::  ; infinity
{ 
    Send("∞")
}

;GREEK
; Lowercase

>^a::  ; alpha
{ 
    Send("α")
} 

>^b::  ; beta
{ 
    Send("β")
} 

>^g::  ; gamma
{ 
    Send("γ")
} 

>^d::  ; delta
{ 
    Send("δ")
} 

>^e::  ; epsilon
{ 
    Send("ε")
} 

>^z::  ; zeta
{ 
    Send("ζ")
} 

>^h::  ; eta
{ 
    Send("η")
} 

>^0::  ; theta  ; unconventional
{ 
   Send("θ")
} 

>^i::  ; iota
{ 
   Send("ι")
} 

>^k::  ; kappa
{ 
   Send("κ")
} 

>^l::  ; lambda
{ 
    Send("λ")
} 

>^m::  ; mu
{ 
    Send("μ")
} 

>^n::  ; nu
{ 
    Send("ν")
} 

>^x::  ; xi
{ 
    Send("ξ")
} 

>^o::  ; omicron
{ 
    Send("ο")
} 

>^p::  ; pi
{ 
    Send("π")
} 

>^r::  ; rho
{ 
    Send("ρ")
} 

>^s::  ; sigma
{ 
    Send("σ")
} 

>^t::  ; tau
{ 
    Send("τ")
} 

>^y::  ; upsilon
{ 
    Send("υ")
} 

>^f::  ; phi
{ 
    Send("φ")
} 

>^c::  ; chi
{ 
    Send("χ")
} 

>^4::  ; psi  ; ($) unconventional
{ 
   Send("ψ")
} 

>^w::  ; omega
{ 
    Send("ω")
}

; Uppercase

>^+g::  ; Gamma
{ 
    Send("Γ")
} 

>^+d::  ; Delta
{ 
    Send("Δ")
} 

; only Rshift
>^>+0::  ; Theta  ; unconventional
{ 
   Send("Θ")
} 

>^+l::  ; Lambda
{ 
   Send("Λ")
} 

; only Rshift
>^>+x::  ; Xi
{ 
  Send("Ξ")
} 

>^+p::  ; Pi
{ 
   Send("Π")
} 

>^+s::  ; Sigma
{ 
  Send("Σ")
} 

; only RShift
>^>+f::  ; Phi
{ 
  Send("Φ")
} 

; only RShift
>^>+4::  ; Psi
{ 
  Send("Ψ")
} 

>^+w::  ; Omega
{ 
  Send("Ω")
} 
