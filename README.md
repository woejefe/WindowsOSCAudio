This program is just a simple OSC listener to control Windows Audio. Written in python and used Pyinstaller to convert to EXE

Currently this is only a Listener and no Polling/Feedback is supported.Wanted to keep it tight and minimize usage on PC.

Seems to only use about 29 MB of RAM consistently and I have noticed next to nothing on CPU usage. 

Default listening port of 9001. 

When you start the program it pulls up a GUI to define listening port, toggle on startup, show logging console, and gives examples. Windows Firewall will prompt you once you start service.

Toggle to start on boot, use tray icon to access GUI. Its a Standalone program so if you want to persistently use you will need to place program in something like "Program Files"


Example OSC:

/master/volume/ 50

/app/volume/firefox/ 72

/mic/mute 1

(only supports default input and output devices currently as well as windows processes)


<img width="523" height="352" alt="OSCAudio" src="https://github.com/user-attachments/assets/0ec9963a-6c7b-4056-85ca-daf94d4e121a" />


Example of a fader in OSCpilot:

<img width="396" height="715" alt="Screenshot 2025-10-29 163033" src="https://github.com/user-attachments/assets/04d9ad28-5c74-463b-ab72-892287f0ee1b" />

-------------------------------------------------------------------------------------------------------------------------------------------



If you want to use python instead of EXE install these dependencies: 
    
    pip install python-osc pycaw comtypes psutil pystray pillow 



This will not work on UNIX systems as it sits. Easy to make a universal app if there is a want.
