icue_display

Converts K95 Platinum Elite Keyboard into a heads up display.

Python Version: 3.9
OS Version: Windows 10

Keyboard HUD Displays:
	* HDD DiskBytesPersec For Each Disk Instance & Each Instance is assigned to its own K95 alpha key.
	* CPU Utilization Monitor
	* DRAM Utilization Monitor
	* VRAM Utilization Monitor

Preferred over Knight Rider C# project becuase:
	* Uses less resources.
	* Self contained control of icue running from threads instead of of running subprograms which could become orphaned.

Frontend is a personal template of mine and currently provides only some control to the backend.

Values are saved to config.dat for memory accross reboots.

Careful not to spam writes to config until readonly boolean's are set while writing, or values spammed will not be written.
