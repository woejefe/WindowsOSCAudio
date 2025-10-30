This program is just a simple OSC listener to control Windows audio.

Default listening port of 9001, when you start program it pulls up GUI to define listening port. Minimizes to Tray, startup option as well.


Example OSC:

/master/volume/ 50

/app/volume/firefox/ 72


<img width="378" height="255" alt="Screenshot 2025-10-29 161620" src="https://github.com/user-attachments/assets/a026178e-cd48-4f31-8730-195729179651" />


Example of a fader in OSCpilot:

<img width="396" height="715" alt="Screenshot 2025-10-29 163033" src="https://github.com/user-attachments/assets/04d9ad28-5c74-463b-ab72-892287f0ee1b" />

-------------------------------------------------------------------------------------------------------------------------------------------



If you want to use python install these dependencies: 
    
    pip install python-osc pycaw comtypes psutil pystray pillow 



