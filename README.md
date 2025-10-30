This program is just a simple OSC listener to control Windows Audio. Written in python and used Pyinstaller to convert to EXE

Currently this is only a Listener and no Polling/Feedback is supported.Wanted to keep it tight and minimize usage on PC.

Seems to only use about 29 MB of RAM consistently and I have noticed next to nothing on CPU usage. 

Default listening port of 9001. 

When you start the program it pulls up GUI to define listening port if youd like to change it. Windows Firewall will prompt you once you start service.

Toggle to start on boot, use tray icon to access GUI. Its a Standalone program so if you want to persistently use you will need to place program in something like "Program Files"


Example OSC:

/master/volume/ 50

/app/volume/firefox/ 72

/mic/mute 1


<img width="378" height="255" alt="Screenshot 2025-10-29 161620" src="https://github.com/user-attachments/assets/a026178e-cd48-4f31-8730-195729179651" />


Example of a fader in OSCpilot:

<img width="396" height="715" alt="Screenshot 2025-10-29 163033" src="https://github.com/user-attachments/assets/04d9ad28-5c74-463b-ab72-892287f0ee1b" />

-------------------------------------------------------------------------------------------------------------------------------------------



If you want to use python instead of EXE install these dependencies: 
    
    pip install python-osc pycaw comtypes psutil pystray pillow 



This will not work on UNIX systems as it sits. Easy to make a universal app if there is a want.
