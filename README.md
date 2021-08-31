# What for is "GPU_Scaling_AspectRatio_Turn_OnOff"?
Should be able to activate GPU-Scaling silent through VBS-Script (Aspect Ratio).

-------------------------------
WARNING - AWARE WHAT YOU DO
-------------------------------
Short form: Use the app at your own risk! And please only use the ready release and NOT the Source!

Long explanation: For your own safety, I strongly recommend to use only the ready EXE. The VBS itself and the third party applications are exactly tuned from each other and available for download as EXE file here.
The source is uploaded here mere on reason of log history. If you use or change the source, it can happen that Windows freezes completely or the graphics card does not output the image anymore, resets your screen settings to default or 'just' messes up the desktop icons! You're warned.
So: Please use only the finished EXE file (relatively stable) and not parts from the source (unsafe).



But now let's focus to that piece of code:
It will change a few special registry entries under HKLM\SYSTEM to apply these changes for AMD Catalyst (maybe for newer versions too, I didn't test it so far). So you'll need admin rights to perform this action. Of course you can do it in AMD's GUI without admin rights but only because there is a service installed it allows you to do so. 
After result, a reboot is required (alternatively execute "Custom Resolution Utility (CRU)" by ToastyX. When AMD loads it's Manager (here ccc.exe), you'll see, GPU-Scaling activated and the mode "Aspect Ratio" is chosen.
Many people said this wouldn't be possible - here I prove they were wrong ;)


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
