# OfficeChromeFix

A slightly kludgy workaround for the fact that the 'flat' UI design of new versions of MS Office make it 
almost impossible to tell which window is active.

I think it also makes life a bit easier if you work with multiple monitors, as even non-Office windows 
in Win7 don't look much different when they're active or inactive.

This program runs in the background and draws a coloured (currently yellow) border around the active window to help you see it.

Before - which window is active?
![Img1](https://raw.githubusercontent.com/michaelomichael/OfficeChromeFix/master/etc/Office1.png)


After - ah, it's the one with the coloured border:
![Img2](https://raw.githubusercontent.com/michaelomichael/OfficeChromeFix/master/etc/Office2.png)


The .exe file to run is: [bin/Debug/WindowFocusNotifier.exe](https://github.com/michaelomichael/OfficeChromeFix/blob/master/bin/Debug/WindowFocusNotifier.exe?raw=true)

There's no way to interact with the program at the moment, so if it's running and you want to kill it, open the Task Manager 
(Ctrl+Shift+Esc), find "WindowFocusNotifier.exe" in the Processes list, and click "End Process".
