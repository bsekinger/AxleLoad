; InnoScript Version 2.3.3
; Randem Systems, Inc.
; Copyright 2003, 2004
; website:  http://www.randem.com
; email:  innoscript@randem.com

; Date: January 30, 2004

;         Visual Basic 6 Runtime Files Folder:   d:\Program Files\Randem Systems\InnoScript\InnoScript 2.3\VB6 Runtime\
;            Visual Basic Project File (.vbp):   D:\My Documents\VB_Progs\Axleload_Unibody\Tread Version\AxleLoad_Unibody_Tread.vbp
;        Inno Setup Script Output File (.iss):   D:\My Documents\VB_Progs\Axleload_Unibody\Tread Version\test.iss
;Visual Basic Project Application File (.exe):   D:\My Documents\VB_Progs\Axleload_Unibody\Tread Version\AxleLoad_Unibody_Tread.exe

; ------------------------
; Visual Basic References
; ------------------------

; Microsoft XML, v3.0
; Windows API (ANSI)
; Microsoft Scripting Runtime


; --------------------------
; Visual Basic Components
; --------------------------

; ComponentOne VSPrinter 8.0 Control
; ComponentOne VSReport 8.0 Control
; ComponentOne VSDraw 8.0 Control
; ComponentOne Sizer/Tab Controls 8.0
; ComponentOne VSFlexGrid 8.0 (Light)
; ComponentOne Chart 8.0 2D Control
; Microsoft Common Dialog Control 6.0 (SP3)
; Microsoft Windows Common Controls 6.0 (SP6)
; Microsoft FlexGrid Control 6.0 (SP3)


[Setup]
AppName=AxleLoad_InHouse
AppVerName=AxleLoad_InHouse 1.0.0.0
DefaultGroupName=Unibody AxleLoad - InHouse Version
AppPublisher=Tread Corporation
;AppPublisherURL=http://www.yourwebsite.com
AppVersion=1.0.0.0
;AppSupportURL=http://www.yourwebsite.com
;AppUpdatesURL=http://www.yourwebsite.com
AllowNoIcons=yes
;InfoBeforeFile=Setup.txt
;InfoAfterFile=ReadMe.txt
;WizardImageFile=yourlogo.bmp
AppCopyright=Copyright 2003 Tread Corporation
PrivilegesRequired=admin
OutputBaseFilename=AxleLoad_Unibody_Tread100
DefaultDirName={pf}\AxleLoad_InHouse

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"
Name: "quicklaunchicon"; Description: "Create a &Quick Launch icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "d:\program files\randem systems\innoscript\innoscript 2.3\vb6 runtime\msvbvm60.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "d:\program files\randem systems\innoscript\innoscript 2.3\vb6 runtime\oleaut32.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "d:\program files\randem systems\innoscript\innoscript 2.3\vb6 runtime\olepro32.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "d:\program files\randem systems\innoscript\innoscript 2.3\vb6 runtime\asycfilt.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  sharedfile
Source: "d:\program files\randem systems\innoscript\innoscript 2.3\vb6 runtime\stdole2.tlb"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  uninsneveruninstall sharedfile regtypelib
Source: "d:\program files\randem systems\innoscript\innoscript 2.3\vb6 runtime\comcat.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "msxml3.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "vsprint8.ocx"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "vsrpt8.ocx"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "vsdraw8.ocx"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "c1sizer.ocx"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "vsflex8l.ocx"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "d:\program files\microsoft visual studio\common\tools\winapi\win.tlb"; DestDir: "d:\program files\microsoft visual studio\common\tools\winapi\"; MinVersion: 4.0,4.0; Flags:  sharedfile regtypelib
Source: "scrrun.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "olch2x8.ocx"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "comdlg32.ocx"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "mscomctl.ocx"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "msflxgrd.ocx"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "d:\my documents\vb_progs\axleload_unibody\tread version\axleload_unibody_tread.exe"; DestDir: "{app}"; MinVersion: 4.0,4.0; Flags:  ignoreversion

[INI]
Filename: "{app}\AxleLoad_Unibody_Tread.url"; Section: "InternetShortcut"; Key: "URL"; String: "http://www.yourwebsite.com"

[Icons]
Name: "{group}\AxleLoad_InHouse"; Filename: "{app}\AxleLoad_Unibody_Tread.exe"; WorkingDir: "{app}"
Name: "{group}\AxleLoad_InHouse on the Web"; Filename: "{app}\AxleLoad_Unibody_Tread.url"
Name: "{group}\Uninstall AxleLoad_InHouse"; Filename: "{uninstallexe}"
Name: "{userdesktop}\AxleLoad_Unibody_Tread"; Filename: "{app}\AxleLoad_Unibody_Tread.exe"; Tasks: desktopicon; WorkingDir: "{app}"

[Run]
Filename: "{app}\AxleLoad_Unibody_Tread.exe"; Description: "Launch AxleLoad_InHouse"; Flags: nowait postinstall skipifsilent; WorkingDir: "{app}"

[UninstallDelete]
Type: files; Name: "{app}\AxleLoad_Unibody_Tread.url"
