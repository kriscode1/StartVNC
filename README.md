# StartVNC
Automatically runs VNC, finds the window of interest, then resizes and repositions VNC appropriately.

This is a time-saving script to automate one of the more tedious tasks of turning on my computer in the morning. 

The code is sufficient for connecting to a VNC server that displays a single window of interest, but the window may resize and reposition from day to day. After connecting through VNC, this script finds that window and centers it to save maximum screen space. Final repositioning coordinates are hardcoded in the script. A future update may remove some of the hardcoded numbers. 

Written for Python 3.5.2
