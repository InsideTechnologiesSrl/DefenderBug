# Defender ASR Bug

On January 13th, Windows Security and Microsoft Defender for Endpoint customers may have experienced a series of false positive detections for the Attack Surface Reduction (ASR) rule "Block Win32 API calls from Office macro" after updating to security intelligence build 1.381.2140.0. These detections resulted in the deletion of certain Windows shortcut (.lnk) files that matched the incorrect detection pattern. 

There’s no way to restore the icon but there’s a way to rebuild the most common to Start Menu with this script.

This script runs on Windows 10 and Windows 11 without elevated permission.

To know what kind of links are deleted, open Microsoft Defender 365 portal and run this Advanced Hunting query:

DeviceEvents
| where Timestamp >= datetime(2023-01-13) and Timestamp < datetime(2023-01-14)
| where ActionType contains "BrowserLaunchedToOpenUrl"
| where RemoteUrl endswith ".lnk"
| where RemoteUrl contains "start menu"
| summarize by Timestamp, DeviceName, DeviceId, RemoteUrl,ActionType
| sort by Timestamp asc

This query is a little bit easy but is without duplicate:

DeviceEvents
| where Timestamp >= datetime(2023-01-13) and Timestamp < datetime(2023-01-14)
| where ActionType contains "BrowserLaunchedToOpenUrl"
| where RemoteUrl endswith ".lnk"
| where RemoteUrl contains "start menu"
| distinct DeviceName, RemoteUrl,ActionType
| sort by RemoteUrl asc

I will keep updated the script every day. If you want collaborate, send your programs.

Last Update: 17/01/2023

Version: 2.1
