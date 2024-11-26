#Requires AutoHotkey v2.0

#SingleInstance Force

#Hotif WinActive("ahk_class Photoshop")

^Numpad9::
{
    quickExportJPGToLastLocation("2160")
}

^Numpad8::
{
    MsgBox("Photoshop is now active")
}

#Hotif ; End photoshop

#Hotif WinActive("ahk_class Premiere Pro")

^!Numpad8::
{
    audioMonoMaker("left")
}

^!Numpad9::
{
    audioMonoMaker("right")
}

;Ripple delete clip at playhead!! This was the first AHK script I ever wrote, I think!
^!Numpad1::
{
    Send "^!s" ;ctrl alt s  is assigned to [select clip at playhead]
    sleep 1
    Send "^+!d" ;ctrl alt shift d  is [ripple delete]
    sleep 1
}

;; instant cut at cursor (UPON KEY RELEASE) -- super useful! even respects snapping!
;note to self, move this to premiere_functions already
;this is NOT suposed to stop the video playing when you use it, but now it does for some reason....
F3::
{
    Send "b" ;This is my Premiere shortcut for the RAZOR tool. You can use another shortcut if you like, but you have to use that shortcut here.
    Send "{shift down}" ;makes the blade tool affect all (unlocked) tracks
    Keywait "F4" ;waits for the key to go UP.
    Send "{lbutton}" ;makes a CUT
    Send "{shift up}"
    Sleep 10
    Send "v" ;This is my Premiere shortcut for the SELECTION tool. again, you can use whatever shortcut you like.
}

;Delete single clip at cursor
^!Numpad3::
{
    prFocus("timeline") ;This will bring focus to the timeline. ; you can't just send ^+!3 because it'll change the sequence if you alkready have the timeline in focus. You have to go to the effect controls first. That is what this function does.
    Send "^!d" ;ctrl alt d is my Premiere shortcut for DESELECT. This shortcut only works if the timeline is in focus, which is why we did that on the previous line!! And you need to deselect all the timeline clips becuase otherwise, those clips will also get deleted later. I think.
    Send "v" ;This is my Premiere shortcut for the SELECTION tool.
    Send "{alt down}"
    Send "{lbutton}"
    Send "{alt up}"
    Send "c" ;I have C assigned to "CLEAR" in Premiere's shortcuts panel.
}

^!Numpad4::
{
    preset("Lumetri SLOG3")
}

#Hotif ; End for Premiere Pro

; Instant App Switching

GroupAdd "yuanjianexploreres", "ahk_class CabinetWClass"

^Numpad1::
{
    if (!WinExist("ahk_class CabinetWClass")) {
        Run "explorer.exe"
    }
    if (WinActive("ahk_exe explorer.exe")) {
        GroupActivate "yuanjianexploreres", "R"
    } else {
        WinActivate "ahk_class CabinetWClass"
    }
}

^Numpad2::
{
    if (!WinExist("ahk_exe chrome.exe")) {
        Run "chrome.exe"
    }
    else if (!WinActive("ahk_exe chrome.exe")) {
        WinActivate "ahk_exe chrome.exe"
    }

}

^Numpad3::
{
    if (!WinExist("ahk_class Premiere Pro")) {
        Run "Adobe Premiere Pro.exe"
    } else if (!WinActivate("ahk_class Premiere Pro")) {
        WinActivate "ahk_class Premiere Pro"
    }

}

^Numpad4::
{
    if (!WinExist("ahk_exe Code.exe")) {
        Run "Code.exe"
    } else if (!WinActivate("ahk_exe Code.exe")) {
        WinActivate "ahk_exe Code.exe"
    }
}

; Adobe Premiere Pro Functions
; ----------------------------
prFocus(panel) {
    ;READ ALL THE COMMENTS BELOW OR SUFFER THE CONSEQUENCES.

    ;PrFocus() allows you to have ONE place where you define your personal shortcuts to "focus" panels in Premiere. It also ensures that panels actually get into focus, and don't rotate through the sequences or anything like that.

    ; THERE IS A FULL VIDEO TUTORIAL THAT TEACHES YOU HOW TO USE THIS, STEP BY STEP.
    ; [[[[[LINK TBD, IT'S NOT FINISHED JUST YET.]]]]]

    ;For this function to work, you MUST Go to Premiere's Keyboard Shortcuts panel, and add the following keyboard shortcuts to the following commands:

    ; KEYBOARD SHORTCUT     PREMIERE COMMAND
    ; ctrl alt shift 3 .....Application > Window > Timeline (The default is shift 3)
    ; ctrl alt shift 1 .....Application > Window > Project  (Sets the focus onto a BIN.) (the default is SHIFT 1)
    ; ctrl alt shift 4 .....Application > Window > Program Monitor (Default is SHIFT 4)
    ; ctrl alt shift 5 .....Application > Window > Effect Controls (Default is SHIFT 5)
    ; ctrl alt shift 7 .....Application > Window > Effects  (NOT the Effect Controls panel!) (Default is SHIFT 7)

    ;(Note that ALL_MULTIPLE_KEYBOARD_ASSIGNMENTS.ahk has the FULL list of keyboard shortcuts that you need to assign in order for my other functions to work.)

    ;EXPLANATION: In Premiere, panel focus is very important. Many shortucts will only work if a specific panel is in focus. That is why those panels must be put into focus FIRST, which is what I built PrFocus() to do. (It's not always as simple as sending just the one keyboard shortcut to activate that panel.)

    ;Right now, we do NOT know which panel is in focus, (or "active.") and AHK has no way to tell (that I know of.)
    ;If a panel is ALREADY in focus, and you send the shortcut to bring it into focus again, that panel might then switch to a different sequence in the case of the timeline or program monitor,, or a different item in the case of the Source panel. IT's a nightmare!

    ;Therefore, we must start with a clean slate. For that, I chose the EFFECTS panel. Sending its focus shortcut multiple times, has no ill effects.

    Sendinput "^!+7" ;bring focus to the effects panel... OR, if any panel had been maximized (using the `~ key by default) this will unmaximize that panel, but sadly, that panel will still be the one in focus.
    ;Note that if the effects panel is ALREADY maximized, then sending the shortcut to switch to it will NOT un-maximize it. This is OK, though, because I never maximize the Effects Panel. If you do, then you might want to switch to the Effect Controls panel first, and THEN the Effects panel after this line.
    ;note that sometimes I use ^+! instead... it makes no difference compared to ^!+

    sleep 12 ;waiting for Premiere to actaully do the above.

    Sendinput "^!+7" ;Bring focus to the effects panel AGAIN. Just in case some panel somewhere was maximized, THIS will now guarantee that the Effects panel is ACTAULLY in focus.

    sleep 5 ;waiting for Premiere to actaully do the above.

    sendinput "{blind}{SC0EC}" ;scan code of an unassigned key. Used for debugging. You do not need this.

    if (panel = "effects")
        goto theEnd ;do nothing. The shortcut has already been sent.
    else if (panel = "timeline")
        Sendinput "^!+3" ;if focus had already been on the timeline, this would have switched to the "next" sequence (in some unpredictable order.)
    else if (panel = "program") ;program monitor. If focus had already been here, this would have switched to the "next" sequence.
        Sendinput "^!+4"
    else if (panel = "source") ;source monitor. If focus had already been here, this would have switched to the next loaded item.
    {
        Sendinput "^!+2"
        ;tippy("send ^!+2") ;tippy() was something I used for debugging. you don't need it.
    }
    else if (panel = "project") ;AKA a "bin" or "folder"
        Sendinput "^!+1"
    else if (panel = "effect controls")
        Sendinput "^!+5"

theEnd:
    sendinput "{blind}{SC0EB}" ;more debugging - you don't need this.
}
;end of prFocus()

preset(item) {
    ;******FUNCTION FOR DIRECTLY APPLYING A PRESET EFFECT TO A CLIP!******
    ; preset() is my most used, and most useful AHK function for Premiere Pro!

    ;===================================================================================
    ; NEW TO AHK? READ ALL THE BELOW INSTRUCTIONS BEFORE YOU TRY TO USE THIS.
    ; THIS WILL NOT WORK UNLESS YOU DO SOME SETUP FIRST!
    ; Fortunately,
    ; THERE IS A FULL VIDEO TUTORIAL THAT TEACHES YOU HOW TO USE THIS, STEP BY STEP.
    ; [[[[[LINK TBD, IT'S NOT FINISHED JUST YET.]]]]]
    ;
    ; Even if Adobe does one day add this feature to Premiere, this video tutorial will
    ; still be very useful to anyone who is learning how to use AHK with Premiere,
    ; especially if you're trying to use any of the other functions that I've created.
    ;
    ; To call the function, use something like
    ; F4::preset("crop 50")
    ; Where "crop 50" is the exact, unique name of the preset you wish to apply.
    ;
    ; For this function to work, you MUST go into Premiere's Keyboard Shortcuts panel,
    ; find the following commands, and add these keyboard shortcut assignments to them:
    ;
    ; Select Find Box ------- CTRL ALT SHIFT F
    ; Shuttle Stop ---------- CTRL ALT SHIFT K
    ; Window > Effects  ----- CTRL ALT SHIFT 7
    ;
    ; (You can use different shortcuts if you like, as long
    ; as those are the ones you send with your AHK script.)
    ;
    ;====================================================================================

    ;Keep in mind, I use 100% UI scaling on my primary monitor, so if you use 125% or 150% or anything else, your pixel distances for commands like Mousemove WILL be different. Therefore, you'll need to "comment in" the message boxes, change some numbers, and keep saving and refreshing and retrying the script until you've got it working!
    ;To find out what UI scaling your screen uses, hit the windows key, type in "display," hit Enter, and then scroll down to "Scale and layout." Under "Change the size of text, apps, and other items," there will be a selection menu thing. Mine is set to "100%." I have NOT done anything in the "Advanced scaling settings" blue link just below that.

    ;To use this script yourself, carefully go through  testing the script and changing the values, ensuring that the script works, one line at a time. use message boxes to check on variables and see where the cursor is. remove those message boxes later when you have it all working!

    ;NOTE: I built this under the assumption that your Effects Panel will be on the same monitor as your timeline. I can't guarantee that it'll work if the Effects Panel is on another monitor.

    ;NOTE: You also need to get the PrFocus() function.

    ;NOTE: To use the preset() function, your cursor must first be hovering over the clip that you wish to apply your preset to. It also works to select multiple clips, again, as long as your cursor is hovering over one of the selected clips.

    ; if (A_PriorHotkey != "") {
    ;     KeyWait A_PriorHotKey ;keywait is quite important.
    ; }
    ;Let's pretend that you called this function using the following line:
    ;F4::preset("crop 50")
    ;In that case, F4 is the prior hotkey, and the script will WAIT until F4 has been physically RELEASED (up) before it will continue.
    ;https://www.autohotkey.com/docs/commands/KeyWait.htm
    ;Using keywait is probably WAY cleaner than allowing the physical key UP event to just happen WHENEVER during the following function, which can disrupt commands like sendinput, and cause cross-talk with modifier keys.

    ;;---------You do not need the stuff BELOW this line.--------------

    SendInput "{blind}{SC0EC}" ;for debugging. YOU DO NOT NEED THIS.
    ;Keyshower(item,"preset") ;YOU DO NOT NEED THIS. -- it simply displays keystrokes on the screen for the sake of tutorials...
    ; if IsFunc("Keyshower")
    ; {
    ; Func := Func("Keyshower")
    ; RetVal := Func.Call(item,"preset")
    ; }
    if !WinActive("ahk_exe Adobe Premiere Pro.exe") ;the exe is more reliable than the class, since it will work even if you're not on the primary Premiere window.
    {
        goto theEnding ;and this line is here just in case the function is called while not inside premiere. In my case, this is because of my secondary keyboards, which aren't usually using #ifwinactive in addition to #if getKeyState(whatever). Don't worry about it.
    }
    ;;---------You do not need the stuff ABOVE this line.--------------

    ;Setting the coordinate mode is really important. This ensures that pixel distances are consistant for everything, everywhere.
    ; https://www.autohotkey.com/docs/commands/CoordMode.htm
    CoordMode("pixel", "Window")
    CoordMode("mouse", "Window")
    CoordMode("Caret", "Window")

    ;This (temporarily) blocks the mouse and keyboard from sending any information, which could interfere with the funcitoning of the script.
    BlockInput "SendAndMouse"
    BlockInput "MouseMove"
    BlockInput "On"
    ;The mouse will be unfrozen at the end of this function. Note that if you do get stuck while debugging this or any other function, CTRL SHIFT ESC will allow you to regain control of the mouse. You can then end the AHK script from the Task Manager.

    SetKeyDelay 0 ;NO DELAY BETWEEN STUFF sent using the "send"command! I thought it might actually be best to put this at "1," but using "0" seems to work perfectly fine.
    ; https://www.autohotkey.com/docs/commands/SetKeyDelay.htm

    SendInput "^!+k" ;in Premiere's shortcuts panel, ASSIGN "shuttle stop" to CTRL ALT SHIFT K.
    sleep 10
    SendInput "^!+k" ; another shortcut for Shuttle Stop. Sometimes, just one is not enough.
    ;so if the video is playing, this will stop it. Othewise, it can mess up the script.
    sleep 5

    ;msgbox, ahk_class =   %class% `nClassNN =     %classNN% `nTitle= %Window%
    ;;This was my debugging to check if there are lingering variables from last time the script was run. You do not need that line.

    MouseGetPos &xposP, &yposP ;------------------stores the cursor's current coordinates at X%xposP% Y%yposP%
    ;KEEP IN MIND that this function should only be called when your cursor is hovering over a clip, or a group of selected clips, on the timeline. That's because the cursor will be returned to that exact location, carrying the desired preset, which it will drop there. MEANING, that this function won't work if you select clips, but don't have the cursor hovering over them.

    SendInput "{mButton}" ;this will MIDDLE CLICK to bring focus to the panel underneath the cursor (which must be the timeline). I forget exactly why, but if you create a nest, and immediately try to apply a preset to it, it doesn't work, because the timeline wasn't in focus...? Or something. IDK.
    sleep 5

    prFocus("effects") ;Brings focus to the effects panel. You must find, then copy/paste the prFocus() function definition into your own .ahk script as well. ALTERNATIVELY, if you don't want to do that, you can delete this line, and "comment in" the 3 lines below:

    ;Sendinput, ^+!7 ;CTRL SHIFT ALT 7 --- In Premiere's Keyboard Shortcuts panel, you nust find the "Effects" panel and assign the shortcut CTRL SHIFT ALT 7 to it. (The default shortcut is SHIFT 7. Because Premiere does allow multiple shortcuts per command, you can keep SHIFT 7 as well, or you can delete it. I have deleted it.)
    ;sleep 12
    ;Sendinput, ^!+7 ;you must send this shortcut again, because there are some edge cases where it may not have worked the first time.

    SendInput "{blind}{SC0ED}" ;for debugging. YOU DO NOT NEED THIS LINE.

    sleep 15 ;"sleep" means the script will wait for 15 milliseconds before the next command. This is done to give Premiere some time to load its own things.

    Sendinput "^!+f" ;SHIFT F ------- set in premiere's shortcuts panel to "select find box"
    sleep 5
    ;Alternatively, it also works to click on the magnifying glass if you wish to select the find box... but this is unnecessary and sloppy.

    ;The Effects panel's find box should now be activated.
    ;If there is text contained inside, it has now been highlighted. There is also a blinking vertical line at the end of any text, which is called the "text insertion point", or "caret".

    caretFound := CaretGetPos(&caretX, &caretY)

    if (!caretFound) {
        ;No Caret (blinking vertical line) can be found.

        ;The following loop is waiting until it sees the caret. THIS IS SUPER IMPORTANT, because Premiere is sometimes quite slow to actually select the find box, and if the function tries to proceed before that has happened, it will fail. This would happen to me about 10% of the time.
        ;Using the loop is also way better than just ALWAYS waiting 60 milliseconds like I was before. With the loop, this function can continue as soon as Premiere is ready.

        ;sleep 60 ;<—Use this line if you don't want to use the loop below. But the loop should work perfectly fine as-is, without any modification from you.

        waiting2 := 0
        loop {
            waiting2 += 1
            sleep 33
            caretFound := CaretGetPos(&caretX, &caretY)
            if (caretFound) {
                Tooltip "counter = " (waiting2 * 33) "`nCaret = " caretX " and " caretY
                break
            }
            if (waiting2 > 40) {
                tooltip("Waiting for too long to find CARET")
                ;Note to self, need much better way to debug this than screwing the user. As it stands, that tooltip will stay there forever.
                ;USER: Running the function again, or reloading the script, will remove that tooltip.
                ;sleep 200
                ;tooltip,
                sleep 20
                goto theEnding
            }
        }
        sleep 1
    }
    ;The loop has now ended.
    ;yeah, I've seen this go all the way up to "8," which is 264 milliseconds

    MouseMove(caretX, caretY, 0) ;this moves the cursor, instantly, to the position of the caret.
    sleep 5 ;waiting while Windows does this. Just in case it takes longer than 0 milliseconds.
    ;;;and fortunately, AHK knows the exact X and Y position of this caret. So therefore, we can find the effects panel find box, no matter what monitor it is on, with 100% consistency!

    MouseGetPos , , &Window, &classNN
    ; winclass := WinGetClass("ahk_id " Window)

    ; tooltip "ahk_class = " winclass "`nClassNN = " classNN "`nTitle= " Window
    ; sleep 2000
    ; tooltip
    ;;;note to self, I think ControlGetPos is not affected by coordmode??  Or at least, it gave me the wrong coordinates if premiere is not fullscreened... IDK. https://autohotkey.com/docs/commands/ControlGetPos.htm

    ControlGetPos &XX, &YY, &Width, &Height, classNN, Window, "SubWindow", "SubWindow"

    ;note to self, I tried to exclude subwindows but I don't think it works...?
    ;;my results:  59, 1229, 252, 21,     Edit1,     ahk_class Premiere Pro
    ;tooltip, classNN = %classNN%

    ;;Now we have found a lot of useful information about this find box. Turns out, we don't need most of it...
    ;;we just need the X and Y coordinates of the "upper left" corner...

    ;;Comment in the following line to get a message box of your current variable values. The script will not advance until you dismiss a message box. (Use the enter key.)
    ; MsgBox "xx= " XX ", yy= " YY

    ;; https://www.autohotkey.com/docs/commands/MouseMove.htm

    ;MouseMove, XX-25, YY+10, 0 ;--------------------for 150% UI scaling, this moves the cursor onto the magnifying glass
    MouseMove XX - 10, YY + 60, 0 ;--------------------for 100% UI scaling, this moves the cursor onto the magnifying glass

    ; MouseMove XX, YY, 0

    ; msgbox "should be in the center of the magnifying glass now." ;;<--comment this in for help with debugging.

    sleep 5

    Sendinput item
    ;This types in the text you wanted to search for, like "crop 50". We can do this because the entire find box (and any included text) was already selected.
    ;Premiere will now display your preset at the top of the list. There is no need to press "enter" to search.

    sleep 5

    ;MouseMove, 62, 95, 0, R ;----------------------(for 150% UI.)
    MouseMove(41, 63, 0, "R") ;----------------------(for 100% UI)
    ;;relative to the position of the magnifying glass (established earlier,) this moves the cursor down and directly onto the preset's icon.

    ;;In my case, all of my presets are contained inside of folders, which themselves are inside the "presets" folder. Your preset's written name should be completely unique so that it is the first and only item.

    ; msgbox "The cursor should be directly on top of the preset's icon. `n If not, the script needs modification."

    sleep 5

    ;;At this point in the function, I used to use the line "MouseClickDrag, Left, , , %xposP%, %yposP%, 0" to drag the preset back onto the clip on the timeline. HOWEVER, because of a Premiere bug (which may or may not still exist) involving the duplicated displaying of single presets (in the wrong positions) I have to click on the Effects panel AGAIN, which will "fix" it, bringing it back to normal.
    ;+++++++ If this bug is ever resolved, then the lines BELOW are no longer necessary.+++++
    MouseGetPos &iconX, &iconY, &Window, &classNN ;---now we have to figure out the ahk_class of the current panel we are on. It might be "DroverLord - Window Class14", but the number changes anytime you move panels around... so i must always obtain the information anew.
    sleep 5
    ; winclass := WinGetClass("ahk_id " Window) ;----------"ahk_id %Window%" is important for SOME REASON. if you delete it, this doesn't work.
    ;tooltip, ahk_class =   %class% `nClassNN =     %classNN% `nTitle= %Window%
    ;sleep 50
    ControlGetPos &xxx, &yyy, &www, &hhh, classNN, Window, "SubWindow", "SubWindow"
    MouseMove(www / 4, hhh / 2, 0, "R") ;-----------------moves to roughly the CENTER of the Effects panel. Clicking here will clear the displayed presets from any duplication errors. VERY important. Without this, the script fails 20% of the time. This is also where the script can go wrong, by trying to do this on the timeline, meaning it didn't get the Effects panel window information as it should have.
    sleep 5
    MouseClick "Left", , , 1 ;-----------------------the actual click
    sleep 5
    MouseMove iconX, iconY, 0 ;--------------------moves cursor BACK onto the preset's icon
    ;tooltip, should be back on the preset's icon
    sleep 5
    ;;+++++If this bug is ever resolved, then the lines ABOVE are no longer necessary.++++++

    MouseClickDrag "Left", iconX, iconY, xposP, yposP, 0 ;---clicks the left button down, drags this effect to the cursor's pervious coordinates and releases the left mouse button, which should be above a clip, on the TIMELINE panel.
    sleep 5
    MouseClick "Middle", , , 1 ;this returns focus to the panel the cursor is hovering above, WITHOUT selecting anything. great! And now timeline shortcuts like JKL will work.

    BlockInput "MouseMoveOff" ;returning mouse movement ability
    BlockInput "Off" ;do not comment out or delete this line -- or you won't regain control of the keyboard!! However, CTRL ALT DELETE will still work if you get stuck!! Cool.

    ;The line below is where all those GOTOs are going to.
theEnding:
}
;END of preset(). The two lines above this one are super important.

audioMonoMaker(track) {
    if !WinActive("ahk_exe Adobe Premiere Pro.exe") {
        goto monoEnding
    }
    Sleep(3)

    CoordMode("Mouse", "Screen")
    CoordMode("Pixel", "Screen")

    BlockInput("On")
    BlockInput("MouseMove")

    global tToggle := 1
    addPixels := 0
    if (track = "right") {
        addPixels := 23 ; right check box is 23 pixels to the right in the Client coordinate system.
    } else if (track = "left") {
        addPixels := 0
    }

    Send("{F3}") ; The audio channels shortcut, default is shift + G in Premiere. Modifier key not reliable in Send().
    Sleep(15)

    MouseGetPos(&xPosAudio, &yPosAudio)
    ; MsgBox("Audio Mouse Pos X, Y: {" xPosAudio ", " yPosAudio "}")
    MouseMove(1129, 833, 0) ; Find the OK button position
    Sleep(15)

    MouseGetPos(&MouseX, &MouseY)
    ; MsgBox("Current Mouse X, Y: {" MouseX ", " MouseY "}")
    waiting := 0

    ; If the mouse is hovering over the OK button, the color should be E8E8E8.
    loop {
        waiting++
        Sleep(50)

        MouseGetPos(&MouseX, &MouseY)
        thecolor := PixelGetColor(MouseX, MouseY)
        ToolTip("waiting = " waiting "`npixel color = " thecolor)
        if (thecolor = "0xE8E8E8") {
            ToolTip("COLOR WAS FOUND")
            break
        }
        if (waiting > 10) {
            ToolTip("no color found, go to ending")
            goto ending
        }
    }

    CoordMode("Mouse", "Client")
    CoordMode("Pixel", "Client")

    MouseMove(110 + addPixels, 201, 0)
    Sleep(50)

    MouseGetPos(&Xkolor, &Ykolor)
    Sleep(50)
    kolor := PixelGetColor(Xkolor, Ykolor)

    if (kolor = "0x1d1d1d" || kolor = "0x333333") {
        MouseClick("left")
        Sleep(10)
    } else if (kolor = "0xb9b9b9") {
        ; Do nothing, checkbox already checked
    }

    Sleep(5)
    MouseMove(110 + addPixels, 224, 0)
    Sleep(30)

    MouseGetPos(&Xkolor2, &Ykolor2)
    Sleep(10)
    k2 := PixelGetColor(Xkolor2, Ykolor2)
    Sleep(30)

    if (k2 = "0x1d1d1d" || k2 = "0x333333") {
        MouseClick("left")
        Sleep(10)
    } else if (k2 = "0xb9b9b9") {
        ; Do nothing, checkbox already checked
    }

    Sleep(5)
    Send("{Enter}")
ending:
    CoordMode("Mouse", "Screen")
    CoordMode("Pixel", "Screen")
    MouseMove(xPosAudio, yPosAudio, 0)
    BlockInput("Off")
    BlockInput("MouseMoveOff")
    ToolTip("")
monoEnding:
}

waitForColor(color, delay := 10) {
    waiting := 0
    loop {
        waiting++
        Sleep(50)

        MouseGetPos(&MouseX, &MouseY)
        thecolor := PixelGetColor(MouseX, MouseY)
        ToolTip("waiting = " waiting "`npixel color = " thecolor)
        if (thecolor = color) {
            ToolTip("COLOR WAS FOUND")
            return true
        }
        if (waiting > delay) {
            ToolTip("no color found, go to ending")
            return false
        }
    }
}

quickExportJPGToLastLocation(width) {
    CoordMode("Mouse", "Screen")
    CoordMode("Pixel", "Screen")

    MouseGetPos(&xPosOrg, &yPosOrg)
    BlockInput("On")
    BlockInput("MouseMove")

    Send("{F10}") ; call out the export as window

    export_x := 1631
    export_y := 898

    MouseMove(export_x, export_y, 0)

    result := waitForColor('0x0D66D0', 100)
    if not result { ; if cannot find the color for a while
        goto ending
    }

    image_width_textbox_x := 1493
    image_width_textbox_y := 319

    MouseMove(image_width_textbox_x, image_width_textbox_y, 0) ; Move to the Width text box
    MouseClick('Left') ; Click the button to get focus
    Sleep(20)
    Send("^a") ; Select All
    Sleep(100)
    Send(width) ; Set the width to be 2160
    Sleep(100)
    MouseClick('Left', 1530, 271) ; Click out of the text box, to apply resolution
    Sleep(20)

    MouseMove(export_x, export_y)
    Sleep(20)
    result := waitForColor('0x0D66D0', 100)
    if not result { ; if cannot find the color for a while
        goto ending
    }

    MouseClickDrag('Left', 1570, 219, 1605, 219, 2) ; Set the JPEG quality of be 7 by draging the slider

    MouseMove(export_x, export_y, 0)

    result := waitForColor('0x0D66D0')
    if not result { ; if cannot find the color for a while
        goto ending
    }

    MouseClick('Left')

    MouseMove(1273, 617) ; The Save button position

    result := waitForColor('0xE0EEF9')
    if not result { ; if cannot find the color for a while
        goto ending
    }

    MouseClick('Left')
ending:
    CoordMode("Mouse", "Screen")
    CoordMode("Pixel", "Screen")
    MouseMove(xPosOrg, yPosOrg, 0)
    BlockInput("Off")
    BlockInput("MouseMoveOff")
    ToolTip("")
}
