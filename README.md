# Create Shortcut
Xojo example project

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

## Description
This example Xojo project to shows to create a ```Shortcut``` *(Windows)*, ```Alias``` *(macOS)* and ```Desktop Launch File``` *(Linux)* in Xojo-built Applications.  
You can use it for example to allow your users to save a Shortcut to launch the application.

### Notes
1. Windows: Using ```OLEObject``` for [Windows Script Host](https://docs.microsoft.com/en-us/troubleshoot/windows-client/admin-development/create-desktop-shortcut-with-wsh)
2. macOS: Declares using ```NSURL```'s [writeBookmarkData](https://developer.apple.com/documentation/foundation/nsurl/1408532-writebookmarkdata?language=objc)
3. Linux: Writes a ```Desktop Launch File```  
   *Note: Depending on the Linux Distribution you need to manually
   right click the created file and choose 'Allow Launching'.*

### ScreenShots
Windows: **Shortcut**  
![ScreenShot: Windows Shortcut](screenshots/windows-shortcut.png?raw=true)

macOS: **Alias**  
![ScreenShot: macOS Alias](screenshots/macos-alias.png?raw=true)

Linux: **Desktop Launch Icon**  
![ScreenShot: Linux Desktop Launch Icon](screenshots/linux-desktop-launch-icon.png?raw=true)

## Xojo
### Requirements
[Xojo](https://www.xojo.com/) is a rapid application development for Desktop, Web, Mobile & Raspberry Pi.  

The Desktop application Xojo example project ```CreateShortcut.xojo_project``` is using:
- Xojo 2023r4
- API 2

### How to use in your own Xojo project?
1. Copy and Paste the Module ```modShortcut``` to your project. Or just copy and paste the extends method ```CreateShortcut``` to one of your global Modules.
2. You can then use it as an extension on a FolderItem:   
    ```anInstanceOfFolderItem.CreateShortcut(saveShortcutFile As FolderItem, useLinuxIconFile As FolderItem = nil)```

## About
Juerg Otter is a long term user of Xojo and working for [CM Informatik AG](https://cmiag.ch/). Their Application [CMI LehrerOffice](https://cmi-bildung.ch/) is a Xojo Design Award Winner 2018. In his leisure time Juerg provides some [bits and pieces for Xojo Developers](https://www.jo-tools.ch/).

### Contact
[![E-Mail](https://img.shields.io/static/v1?style=social&label=E-Mail&message=xojo@jo-tools.ch)](mailto:xojo@jo-tools.ch)
&emsp;&emsp;
[![Follow on Facebook](https://img.shields.io/static/v1?style=social&logo=facebook&label=Facebook&message=juerg.otter)](https://www.facebook.com/juerg.otter)
&emsp;&emsp;
[![Follow on Twitter](https://img.shields.io/twitter/follow/juergotter?style=social)](https://twitter.com/juergotter)

### Donation
Do you like this project? Does it help you? Has it saved you time and money?  
You're welcome - it's free... If you want to say thanks I'd appreciate a [message](mailto:xojo@jo-tools.ch) or a small [donation via PayPal](https://paypal.me/jotools).  

[![PayPal Dontation to jotools](https://img.shields.io/static/v1?style=social&logo=paypal&label=PayPal&message=jotools)](https://paypal.me/jotools)
