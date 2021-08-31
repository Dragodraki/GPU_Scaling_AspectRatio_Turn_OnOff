# What for is "GPU_Scaling_AspectRatio_Turn_OnOff"?
Should be able to toggle GPU-Scaling (Aspect-Ratio vs. Stretch) no mind whether you use a GPU from Nvidia, AMD or Intel and independent of whether full driver manager software is installed or only the driver.

Many people said this wouldn't be possible - here I am to prove they were wrong ;)

-------------------------------
WARNING - AWARE WHAT YOU DO
-------------------------------
Short form:
- Use the app at your own risk!
- Please only work with the ready release and NOT the Source (may crash your system and cause data loss)!
- Don't re-upload with your name as programmer.
- Don't sell it as product!
+ Commcercial usage is allowed ;)
+ Free Distribution to friends or strangers is allowed, even wanted ;)


Long explanation: 

For your own safety, I strongly recommend to use only the ready EXE. The VBS itself and the third party applications are exactly tuned from each other and available for download as EXE file here.
The source is uploaded here mere on reason of log history. If you use or change the source, it can happen that Windows freezes completely or the graphics card does not output the image anymore, resets your screen settings to default or 'just' messes up the desktop icons! You're warned.
So: Please use only the finished EXE file (relatively stable) and not parts from the source (unsafe).

But now let's focus to that piece of code:
It will change a few special registry entries under HKLM\SYSTEM and HKLM\SOFTWARE to apply these changes for various GPU manufacturer drivers. So you'll need admin rights to perform this action. Of course you can do it in the Graphiccard GUI manager without admin rights but only because there is a service installed it allows you to do so. 
After result, a reboot is required to take settings effect. But don't panic - there's a workaround fo this implemented via "Custom Resolution Utility (CRU)" by ToastyX. It restart all available graphic cards automatically. If you've a GPU GUI manager installed, you should see GPU-Scaling status changed and the mode "Aspect Ratio" / "Full screen" is chosen.

One word please on the topic of distribution, rights and usage: My tool includes the VBS-script as it's main feature but need the help of third-party-apps like "DesktopOK", "Custom Resolution Utility (CRU)" and "PScript". They won't be listed in my Source folder because I'm not the inventor. Although my part is freeware it doesn't mean you are allowed to either sell the app (or parts of it) itself or repacking it and distribute it in your name - additionally does the terms of the third-party-apps applies to!
ANY DISBEHAVIOUR AGAINST THESE RESTRICTIONS I ASSUME NO LIABILITY!

-------------------------------
WHY I MADE THIS SCRIPT?
-------------------------------
I tried to find a solution to make old games fitting the screen perfectly on new widescreen monitors without manually change it in AMD/Nvidia/Intel.
For future I'd like to improve the code so thats work for mostly every great graphic manufacturer software.

-------------------------------
HOW IT WORKS TECHNICALLY:
-------------------------------
The settings for the screen are managed by Windows as well as by drivers of the graphics card. The selected mode of enabled scaling (aspect-ratio/centered/full-screen) is always controlled by Windows without exception - this setting is always stored in the same place in the registry.
The setting whether GPU scaling is enabled or disabled (GPU scaling versus display scaling) is up to the graphics card software or driver. Depending on whether the graphics card is from Nvidia/AMD/Intel and which driver/control center version is installed, this setting is located in a different place in the registry each time, which is not permanently defined.

My software does practically nothing else than to predefine the entries for scaling at the various places in the registry for the most popular graphics card manufacturers, so that the driver reads this entry and loads the appropriate setting - hence the reloading of the graphics card becomes so important.
And this means that my app is not able to replace your graphics card manufacturer's driver, without GPU driver installed my app is useless. Nevertheless, finding out the places in the registry that are responsible for scaling was very difficult, especially since the AMD/Nvidia/Intel didn't want to help me develop this background app. So consider it a small milestone in screen history that it now works without the oh-so-great interface software from the big-name manufacturers.

-------------------------------
SPECIAL THANKS
-------------------------------
I'm so glad this project become so much smoothier because of your tools. Thank you Jordan Russell, Martijn Laan, ToastyX, Nenad and Robert D. Riemersma, Jr! Without your help my VBS script never had a chance to be executed as standalone EXE and without: annoying reboots, pitiful icon rearrangements and regretting in using microsofts version microsofts vbs scripting host.

-------------------------------
SUPPORTED GPUs
-------------------------------
Note: Proper gpu driver needs to be installed.

- ( work ) AMD Catalyst Control (as long as a newer driver is used for it, this old interface does not offer scalable drivers itself
- ( work ) AMD Radeon
- ( work ) AMD Adrenalin
- ( work ) Intel Graphics Control Panel
- ( work ) Nvidia Control Panel
- ( ????? ) Nvidia GeForce Experience

-------------------------------
USAGE
-------------------------------
1. Download the release (or the EXE files from branch) and execute it (need admin rights).
2. Ready - be happy about fixed aspect ratio :D
