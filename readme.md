# AutoHotKey Utilities

## Credits

- [AHK forum post by lexikos](https://www.autohotkey.com/boards/viewtopic.php?t=9832) for middle click to open new folder
- [Stack Overflow answer by Ethan Malloy](https://stackoverflow.com/a/75881915) for toggle current window always-on-top
- formatc1702's [KeyboardMacros](https://github.com/formatc1702/KeyboardMacros), which this is forked from
- [AHK-v2-script-converter](https://github.com/mmikeww/AHK-v2-script-converter)

## Keys

### Misc

|||
|------|------|
|Right Alt + ,|**Browser Back**|
|Right Alt + .|**Browser Forward**|
|Right Alt + o|**Open OneDrive (or another folder)¹**|
|Shift + Mouse Wheel|**Horizontal Scroll²**|
|Middle Mouse Button (in File Explorer)|**Open folder in new window**|
|Right Alt + L (in Word)|**Switch to one page view³**|
|Right Control + Enter|**Toggle current window always-on-top**|
|Right Control + Backspace|**Alternative superscript hotkey in some Office apps⁴**|
|Right Control + \ |**Alternative subscript hotkey in some Office apps⁴**|

¹Edit the script with your OneDrive/folder path to use.

²Many programs don't need this, but it is useful in OneNote. Excel is still Control + Shift + Mouse Wheel.

³E.g., when zooming out.

⁴This is commented out, but depending on preference I reccomend uncommenting it back in. The programs it affects are Word, PowerPoint, OneNote, and Outlook. 

- Note that currently the stock sub/superscript hotkeys in Word are different from in PowerPoint/OneNote/Outlook, which is implicated in how this hotkey and the subscript hotkeys are written. If the other three stock kotkeys are changed, this script will hopefully be updated with a fix.

### Greek Letters

#### **Lower Case Greek Letters:** Left Control + Key

|Key||||*(...)*|||
|---|------|-------|-----------|------|-------|-------|
|a  |α     |**Alpha**  |           |n     |ν      |**Nu**     |
|b  |β     |**Beta** |           |x     |ξ      |**Xi**     |
|g  |γ     |**Gamma**  |           |o     |ο      |**Omikron**|
|d  |δ     |**Delta**  |           |p     |π      |**Pi**     |
|e  |ε     |**Epsilon**|           |r     |ρ      |**Rho**    |
|z  |ζ     |**Zeta**   |           |s     |σ      |**Sigma**  |
|h  |η     |**Eta**    |           |t     |τ      |**Tau**    |
|0  |θ     |**Theta**  |           |y     |υ      |**Upsilon**|
|i  |ι     |**Iota**   |           |f     |φ      |**Phi**    |
|k  |κ     |**Kappa**  |           |c     |χ      |**Hi**     |
|L  |λ     |**Lambda** |           |4     |ψ      |**Psi**    |
|m  |μ     |**Mu**     |           |w     |ω      |**Omega**  |

#### **Upper Case Greek Letters:** Left Control + Shift + Key

|Key|Right Shift only?|||
|------|-----------------|------|------|
|g     |                 |Γ     |**Gamma** |
|d     |                 |Δ     |**Delta** |
|0     |RS              |Θ     |**Theta**|
|L     |                 |Λ     |**Lambda**|
|x     |RS              |Ξ     |**Xi**   |
|p     |                 |Π     |**Pi**    |
|s     |                 |Σ     |**Sigma** |
|f     |RS              |Φ     |**Phi**  |
|4     |RS              |Ψ     |**Psi**  |
|w     |                 |Ω     |**Omega** |

### Math and other symbols

"LS", "RC", etc. indicate use of specifically left or right Control, Shift, or Alt.

#### **General**

|Key|Control|Shift|Alt||                |
|---|-------|-----|---|------|----------------------|
|v  |RC     |     |   |√     |**Square root**           |
|v  |RC     |S    |   |✓     |**Check mark**            |
|* (Numpad)|       |RS   |   |°     |**Degree**                |
|6  |RC     |     |   |°     |**Degree**                |
|`  |RC     |     |   |°     |**Degree**                |
|`  |RC     |S    |   |•     |**Bullet**                |
|'  |RC     |     |   |•     |**Bullet**                |
|Backspace|RC     |S    |   |≈     |**Approximately equal to**|
|\  |RC     |S    |   |±     |**Plus or minus**         |
|8  |RC     |     |   |∞     |**Infinity**              |
|2  |RC     |     |A  |½     |**Half**                  |

#### **Subscripts**

⭐ In Word, PowerPoint, OneNote, and Outlook, superscripts and subscripts are inserted as if they were typed manually. Elsewhere, their Unicode characters are used.

|Key|Control|Shift|Alt||                |
|---|-------|-----|---|------|----------------------|
|0  |RC     |LS   |   |X₀    |**Subscript zero**        |
|1  |RC     |S    |   |X₁    |**Subscript one**         |
|2  |RC     |S    |   |X₂    |**Subscript two**         |
|3  |RC     |S    |   |X₃    |**Subscript three**       |
|4  |RC     |LS   |   |X₄    |**Subscript four**        |
|5  |RC     |S    |   |X₅    |**Subscript five**        |
|6  |RC     |S    |   |X₆    |**Subscript six**         |
|7  |RC     |S    |   |X₇    |**Subscript seven**       |
|8  |RC     |S    |   |X₈    |**Subscript eight**       |
|9  |RC     |S    |   |X₉    |**Subscript nine**        |
|n  |RC     |S    |   |Xₙ    |**Subscript n**           |
|i  |RC     |S    |   |Xᵢ    |**Subscript i**           |
|f  |RC     |LS |   |X<sub>f</sub>    |**Subscript f*** |
|x  |RC     |LS   |   |Xₓ    |**Subscript x**           |

*No Unicode character exists for subscript f, so this only works in the four Office apps.

#### **Superscripts**

|Key|Control|Shift|Alt||                |
|---|-------|-----|---|------|----------------------|
| -  |RC     |     |   |X⁻    |**Superscript minus**       |
|=  |RC     |     |   |X⁺    |**Superscript plus**        |
|0  |RC     |     |A  |X⁰    |**Superscript zero**      |
|1  |RC     |     |   |X¹    |**Superscript one**       |
|2  |RC     |     |   |X²    |**Superscript two**       |
|3  |RC     |     |   |X³    |**Superscript three**     |
|4  |RC     |     |A  |X⁴    |**Superscript four**      |
|5  |RC     |     |A  |X⁵    |**Superscript five**      |
|6  |RC     |     |A  |X⁶    |**Superscript six**       |
|7  |RC     |     |A  |X⁷    |**Superscript seven**     |
|8  |RC     |     |A  |X⁸    |**Superscript eight**     |
|9  |RC     |     |A  |X⁹    |**Superscript nine**      |
|n  |RC     |     |A  |Xⁿ    |**Superscript n**         |

## Installation on Windows

This repo includes both AHK V1.1 and V2 versions of this script. If you do not have either installed, I recommend first installing V2. If you later come across a V1.1 script to run, AHK will allow you to automatically download and install V1.1 alongside your V2 install.

1. Download and install [AutoHotKey](https://www.autohotkey.com/).
2. Download either all_V1.1.ahk or all_V2.ahk, depending on your version of AHK.
3. Right click the script in Explorer and select edit to make any desired changes. If you would like to run the script now, double click the script. To stop it, right click it in the system tray and select Exit.
4. To have the script lauch at startup, paste a shortcut to the script in the Startup folder 
(%appdata%\Microsoft\Windows\Start Menu\Programs\Startup). You can get to this folder quickly by running Win + R, entering <kbd>shell:startup</kbd>, and then pressing Enter.

The [AHK docs](https://www.autohotkey.com/docs/) are helpful for troubleshooting.

## General notes

If the script doesn't launch fast enough at startup using the method listed, you can also put the script in a folder (along with any other AHK scripts you would like) and create a batch file named startahk.bat. Edit the file to include the following:

```
@echo off
for %%a in ("E:\OtherSoftware\AHKStart\*.ahk") do start "" "%%a"
```

Then, open the Windows Task Scheduler. On the right, click Create Task. Name it AHKStartup or similar. Under Triggers, select New and choose to begin the task at logon (of either any user or just your user). In Conditions, uncheck "Start the task only if the computer is on AC power" (for laptops). Under Actions, select New and browse for startahk.bat (or paste its path). Then click OK twice.
