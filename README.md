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
