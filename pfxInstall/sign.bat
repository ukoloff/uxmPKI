Set EXE=pfxInstall.exe
Set TOOL="C:\Program Files\Microsoft SDKs\Windows\v7.1\Bin\signtool.exe" 
Set TS=http://timestamp.verisign.com/scripts/timstamp.dll
upx %EXE%
%TOOL% sign %EXE%
%TOOL% timestamp /t %TS% %EXE%
