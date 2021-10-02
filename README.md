# What for is "GPU_Scaling_AspectRatio_Turn_OnOff"?
Should be able to toggle GPU-Scaling (Aspect-Ratio vs. Stretch) no mind whether you use a GPU from Nvidia, AMD or Intel and independent of whether full driver manager software is installed or the driver only.

Many people said this wouldn't be possible - here I am to prove they were wrong ;)

<img src="https://user-images.githubusercontent.com/76787321/135721075-d583fd94-b5d1-489b-8d3f-828d33792cf2.png" width="15%"></img> 

-------------------------------
WARNING - AWARE WHAT YOU DO
-------------------------------
Short form:
- Use the app at your own risk!
- Please only work with the latest release and NOT older releases NOR the Source (may crash your system and cause data loss - I'm serious)!
- Don't re-upload with your name as programmer.
- Don't sell it as product!
+ Commcercial usage is allowed ;)
+ Free Distribution to friends or strangers is allowed, even wanted ;)


Long explanation: 

For your own safety, I strongly recommend to use only the ready EXE of the latest release. The VBS itself and the third party applications are exactly tuned from each other and available for download as EXE file here.
The source is uploaded here mere on reason of log history. If you use or change the source, it can happen that Windows freezes completely or the graphics card does not output the image anymore, resets your screen settings to default or 'just' messes up the desktop icons! Even when using the release in rarely cases leads to frozen start menu in Windows 10+ last until reboot (since version 1.2 largely fixed) and crashed VMware virtual machines (since version 1.3 largely fixed). You're warned.

But now let's focus to that piece of code:
It will change several registry entries under HKLM\SYSTEM and HKLM\SOFTWARE to apply these changes for various GPU manufacturer drivers. So you'll need admin rights to perform this action. Of course you can do it in the graphic card GUI manager without admin rights but only because there is a service installed it allows you to do so. 
After result, a reboot is required to take settings effect. But don't panic - there's a workaround fo this implemented via "Custom Resolution Utility (CRU)" by ToastyX. It restart all available graphic cards automatically. If you've a GPU GUI manager installed, you should see GPU-Scaling status changed and the mode "Aspect Ratio" / "Full screen" is chosen.

One word please on the topic of distribution, rights and usage: My tool includes the VBS-script as it's main feature but need the help of third-party-apps like "DesktopOK", "Custom Resolution Utility (CRU)", "PScript" and "SetACL". They won't be listed in my Source folder because I'm not the inventor. Although my part is freeware it doesn't mean you are allowed to either sell the app (or parts of it) itself or repacking it and distribute it in your name - additionally does the terms of the third-party-apps applies to!

ANY DISBEHAVIOUR AGAINST THESE RESTRICTIONS OR DAMAGE TO YOUR SYSTEM BY MY SOFTWARE I ASSUME NO LIABILITY!

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

- My software does practically nothing else than to predefine the entries for scaling at the various places in the registry for the most popular graphics card manufacturers, so that the driver reads this entry and loads the appropriate setting - hence the reloading of the graphics card becomes so important.
And this means that my app is NOT able to replace your graphics card manufacturer's driver, without GPU driver installed my app is useless. Nevertheless, finding out the places in the registry that are responsible for scaling was very difficult, especially since the AMD/Nvidia/Intel didn't want to help me develop this background app. So consider it a small milestone in screen history that it now works without the oh-so-great interface software from the big-name manufacturers.

Additionally there are a few tools that saves you from suffering from unwanted side effects:
- "DesktopOK" saves the position of all desktop icons before reloading GPU and restoring after that (I don't rely on Windows since some graphic cards infere with Microsoft and the consequence would have been all icons placed on the left side and had to be arranged manually - bah! Thanks to teams from DesktopOK we are ahead of this.
- "SetACL" helps to unlock denied registry writing permissions. Why? ...I would like to know that too. But as long as TrustedInstaller keep blocking by default half of the (less-important) registry part, we need to unlock the section for gpu display entries first before editing them. Without, settings wouldn't be applied.
- "PScript" is an alternative for the microsoft scripting host (player for VBS files). It works great on newer operating systems like Windows Vista and higher. FOr Windows 2000 and XP the old wscript.exe from microsoft will be used.
- "Custom Resolution Utility (CRU)" - yeah - thats mostly the key function for the undeniable advantage that your computer doesn't need to reboot after changing your graphic cards driver settings. Of course, AMD, Intel and Nvidia could bypass them easily as they are the inventor od their driver and know some seret approaches to them we'll never get to explained. But CRU has the power to adapt it's behaviour in disable and enable your graphic card in device manager from command line. But instead of manual doing they do it a little safer so there is no risk the display won't awake from blackness after this. For your information: Until now, no other free tool exists that can perform that!
- "Rexplorer" preempts problems related to Start menu/app windows/taskbar by killing all processes called "explorer.exe" and restart them with restoring all instances of Windows-Explorer / system control. On all OS before Windows 8 it interrupts also any file copying. How lucky we must be, start menu freezing only happends on Windows 10, so Rexplorer don't harm any open copy process (it only closes the progress window). On older OS my App don't use Rexplorer.
- "vmrun" from VMware shouldn't named here actually since I have to use it only for taking prevention of a bug by VMware itself leads to heavy crashing it's VM's when running ToastyX "restart.exe" (reloads the GPU). But the software is very powerful when it comes to handle the VM's silently. Sorry about the short console window appearing, but to make it invisible would costs disproportionate too much effort for me right now (maybe in the future).

-------------------------------
SPECIAL THANKS
-------------------------------
I'm so glad this project become so much smoothier because of your tools. Thank you Jordan Russell, Martijn Laan, ToastyX, Nenad, Robert D., Riemersma Jr, Helge Klein and Sordum team! Without your helpful tools my VBS script never had a chance to be executed as standalone EXE and without: annoying reboots, pitiful icon rearrangements and regretting in using microsofts version vbs scripting host as like as unnecessary denied registry entries.

-------------------------------
SUPPORTED GPUs
-------------------------------
Note: Proper gpu driver needs to be installed. And OS older than Windows Vista haven't prerequisites for scaling with this tool.

- ( work ) AMD Catalyst Control (as long as a newer driver is used for it, this old interface does not offer scalable drivers itself
- ( work ) AMD Radeon
- ( work ) AMD Adrenalin
- ( work ) Intel Graphics Control Panel (old interface)
- ( work ) Intel Graphics Control Panel (new interface)
- ( work ) Nvidia Control Panel (old interface)
- ( work ) Nvidia Control Panel (new interface - at least until 2021)

-------------------------------
USAGE
-------------------------------
1. Download the release (or the EXE files from branch) and execute it (need admin rights).
2. Ready - be happy about fixed aspect ratio :D
