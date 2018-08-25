# fsBacnetWrite
the program to write a value to multiple Bacnet points (in different devices)

Release bin file is inside "Release.zip"

Unzip the "Release.zip" and copy the VAV0_list.xlsx to same folder with "fsBacnetWrite.exe" (inside Release folder)

To write 10.0 to all points with Bacnet IP of 10.2.22.2:
```
fsBacnetWrite.exe --bacnetip 10.2.22.2 --value 10.0
```
To write null to all points with Bacnet IP of 10.2.22.2:
```
fsBacnetWrite.exe --bacnetip 10.2.22.2
```

Source code is inside "Program.fs"

To build:
```
dotnet restore
dotnet build
```

## Discovery TimeOut
The default time for WhoIs process to discover devices on the system is 1s. On the system with many devices, it might be needed to increase this time so that all the devices are listed. This can be done by ```--timeout``` argument.
For example, to use 2s for timeout the command is
```
fsBacnetWrite.exe --bacnetip 10.2.22.2 --timeout 2.0
```
