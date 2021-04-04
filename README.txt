iCUE Display

Converts K95 Platinum Elite Keyboard into a heads up display. Providing data and a type of security whereby
one can see what is happening on the system.


Keyboard HUD Displays:
	* HDD DiskBytesPersec For Each Disk Instance & Each Instance is assigned to its own K95 alpha key.
	* CPU Utilization Monitor
	* DRAM Utilization Monitor
	* VRAM Utilization Monitor

GUI:
Designed to be purely logical & non-resource heavy to be used as a tool.

Enable/Disable: - Turns functions on/off.
RGB 1: Colour of RGB on.
RGB 2: Colour of RGB off.

Timing: Accepts value between 0.1 and 5. Lower the number, more respources iCUE Display will consume to display information
to the keyboard

Network Traffic:
1-9 Blue
10-99 Light Blue
100+ White
Bytes Red, KB Green, MB Blue, GB Light Blue, TB White
SENT: Intuitively furthest away. F1 - F9. F10 Being unit/ten/hundred
REVEIVED: Intuitively closest. 1 - 9. 0 Being unit/ten/hundred

Recommended Settings:
Timing: All 1.0 except HDD & Network traffic at 0.1
Exclusive Connection(Exclusive Control): Enabled
RGB 2: 0,0,0 (off) for all.


Install:
iCUEDisplay.exe will create its files and shortcut withing the directory iCUEDisplay is ran. iCUE Display can then be ran
quietly from the created shortcut.

Requirements:
Python Version: 3.9
OS Version: Windows 10
ICUE Running

Note: This program is not a lighting display effect. It is designed to convert the icue keyboard into a non-distracting,
efficient and non-resource heavy HUD.

Developer notes:
iCUE Display executable is compiled from .py so that module GPUINFO can pipe stdout which is necessary for gpu information.
(This is why the vbs to run quietly will be used because iCUE display is compiled from .py rather than .pyw).

iCUE Display ius designed to be efficient. Threads are killed when iCUE is not running. Threads are killed when function(s)
disabled. Use recommended settings for best efficiency and use.

Shortcuts created are intended for use with iCUEDisplay.exe. Otherwise you must have python and required modules to run
iCUEDisplay.py.
