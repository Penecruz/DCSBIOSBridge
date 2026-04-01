# DCSBIOSBridge

I have forked this repository for my own cockpit building purposes, mainly to update some featues and add some resiience for my Simpit connections. Feel free to use it at your own risk, however no support will be provided. I'm happy to accept any issues but offer no guarentee of a quick fix.

 ## DCS-BIOS Bridge handles communication between DCS-BIOS and serial ports.

This Version incudes custom names for Comm ports to refect the attached DCS-BIOS Device, the option to open all comm ports on DCS-BIOS Bridge startup, a Watch Dog and display counter that monitors the number of times it barks to wake a non responsive device after a read or write from the connected device below a configurable threshold value, This will help stop a device becoming unresponsive in game (it will probably NOT solve comm ports dropping out due to hardware configuration issues).

V1.2.1 
- Retain Custom names on COM Port disconnection and reconnect.
- If a Com port recconcts and Open all ports on startup is checked, port will open on reconnect.
- Added windows themee options Follow Windows, Light and Dark modes. 

Many thanks to the guys at DCS-Skunkworks who kept DCS-BIOS alive!
 
<img width="586" height="504" alt="image" src="https://github.com/user-attachments/assets/f5b29928-29b9-4eb0-b8de-587a250c29fc" />

<img width="515" height="685" alt="image" src="https://github.com/user-attachments/assets/bae56c85-6784-4fff-824a-40181f8c1954" />

New configuration options for Opening all ports on startup, Theme, and also gain controls for the Watch Dog's aggresivness.

<img width="284" height="838" alt="image" src="https://github.com/user-attachments/assets/b20a957a-368c-46db-a75e-4e2176e43890" />

A little improvement to the About window with some more user information.

<img width="1029" height="1522" alt="image" src="https://github.com/user-attachments/assets/d3f45b86-5947-4ac0-a3ed-86a9bf4843d8" />


As I make modifications for my own simpit and DCS-BIOS Bridge I will keep this repo updated with the latest additions and enhancements.

Cheers PeneCruz


