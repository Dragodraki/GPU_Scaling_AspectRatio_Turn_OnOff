# What for is "GPU_Scaling_AspectRatio_Turn_OnOff"?
Should be able to toggle GPU-Scaling (Aspect-Ratio vs. Stretch) no mind whether you use a GPU from Nvidia, AMD or Intel and independent of whether full driver manager software is installed or only the driver.

Many people said this wouldn't be possible - here I am to prove they were wrong ;)

-------------------------------
WARNING - AWARE WHAT YOU DO
-------------------------------
Short form: Use the app at your own risk! And please only use the ready release and NOT the Source!

Long explanation: For your own safety, I strongly recommend to use only the ready EXE. The VBS itself and the third party applications are exactly tuned from each other and available for download as EXE file here.
The source is uploaded here mere on reason of log history. If you use or change the source, it can happen that Windows freezes completely or the graphics card does not output the image anymore, resets your screen settings to default or 'just' messes up the desktop icons! You're warned.
So: Please use only the finished EXE file (relatively stable) and not parts from the source (unsafe).

But now let's focus to that piece of code:
It will change a few special registry entries under HKLM\SYSTEM and HKLM\SOFTWARE to apply these changes for various GPU manufacturer drivers. So you'll need admin rights to perform this action. Of course you can do it in the Graphiccard GUI manager without admin rights but only because there is a service installed it allows you to do so. 
After result, a reboot is required to take settings effect. But don't panic - there's a workaround fo this implemented via "Custom Resolution Utility (CRU)" by ToastyX. It restart all available graphic cards automatically. If you've a GPU GUI manager installed, you should see GPU-Scaling status changed and the mode "Aspect Ratio" / "Full screen" is chosen.


-------------------------------
WHY I MADE THIS SCRIPT?
-------------------------------
I tried to find a solution to make old games fitting the screen perfectly on new widescreen monitors without manually change it in AMD/Nvidia/Intel.
For future I'd like to optimize the code so thats work for mostly every great graphic manufacturer software.


-------------------------------
USAGE
-------------------------------
1. Optional (but recommended): Copy "wscript.exe" from C:\Windows\System32 in one of your profile folder and mark under compatibillity options the "Run As Administrator".
2. Drag the VBS and drop it over your new "wscript.exe" from step 1 and apply UAC (if UAC is enabled)
3. Script is ready within a few millisecconds. You won't see any difference, because it works silent. You have to reboot your computer or execute "Custom Resolution Utility (CRU)" by ToastyX to settings take effect.
4. Run an old game with e.g. 4:3 or 3:2 scale and see it working.
