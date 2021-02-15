# UpdateConexantISSTDriver

C# program to update the Conexant ISST Driver via Windows Update. 

This program was created to help update Windows 10 operating systems with the latest Conexant ISST drivers available for a device on Windows Update. 
Many devices have been stopped from upgrading to Windows 10 20H1 + by Safeguard Hold ID 25178825 where an outdated Conexant ISST driver is the cause.

It is intended for corporate automated deployment with tools like Microsoft Endpoint Configuration Manager.

This program must run with administrative privege or in SYSTEM context and does the following:
- Checks if registry keys have been set which prevent access to Windows Update online (eg local WSUS server) and temporarily opens up access if necessary
- Uses the Windows Update agent to check if a new Conexant ISST driver is available for the device. These start with "Conexant - MEDIA" in the update catalog.
- Downloads and installs the driver
- Restores the previous Windows Update access settings if necessary
- Logs to the Temp directory, eg C:\Windows\Temp\ConexantDriverUpdate.log for the SYSTEM context

A reboot is usually required and the driver will usually display a small one-time toast notification advising the user to restart.

You can clone the solution in Visual Studio and build your own executable or download the package containing it.
