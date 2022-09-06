# symmetrical-waffle
VB6 ScreenShots, and other software,  from TDS5054 and TDS3000 Scopes
 '**********************************************************************************************************************************************************
 '******** Written by David Kuhn of Klassic Benchmark NDE Repair and Calibration Servicess, LLC., September 2022                                    ********
 '******** klassicbenchmark@gmail.com                                                                                                               ********
 '********                                                                                                                                          ********
 '******** KB Release Revision 1.0.0                                                                                                                ********
 '********                                                                                                                                          ********
 '******** This code is free to use as long as credit is given.                                                                                     ********
 '********                                                                                                                                          ********
 '******** Currently it is supporting the TDS3000 series scopes and the TDS5054B (Tested with a TDS5054B-NV-AV)                                     ********
 '********                                                                                                                                          ********
 '******** This was written and tested specifically with VB6, but should work with VBA (Microsoft Offic)                                            ********
 '********                                                                                                                                          ********
 '******** Yor program must have GLOBmgr.DLL references as the VISA interface to GPIB/LAN/USB.                                                      ********
 '******** GPIB Addess is the full address (i.e., GPIB0::1::INSTR).                                                                                 ********
 '********                                                                                                                                          ********
 '******** The optional variables should be self-explanatory.  FileName is specifically for the TDS5054 and is the name created on that scope's     ********
 '******** internal hard drive. That file is deleted after the file is downloaded from the GPIB port.                                               ********
 '********                                                                                                                                          ********
 '******** If anyone improves this, by all means, please email me the updated version.  I am by no means an expert coder.                           ********
 '********                                                                                                                                          ********
 '**********************************************************************************************************************************************************

Function Capture_ScopeScreen(gpib_address As String, file_path As String, FileType As Integer, InkSaver As Boolean, Optional IsTDS5054 As Boolean, Optional FileName As String, Optional JustGrat As Boolean) As Boolean
On Error GoTo ProcError
'
Dim sBuffer As String, Ext As String, SaveInk As String, FileFormat As String, TDS5000SaveInk As String, Palette As String, ScreenShow As String
Dim data_array As Variant, xvalue As Variant, yvalue As Variant
Dim i As Integer, fn As Integer
Dim FileSize As Long
Dim byteData() As Byte
Dim LioMgr As VisaComLib.ResourceManager
Dim equip As VisaComLib.FormattedIO488
'
Capture_ScopeScreen = True
'
'check for missing arguments
If gpib_address = "" Then
    Err.Raise (448)
End If
'
If Left(gpib_address, 4) <> "GPIB" Then
    Err.Raise (380)
End If
'
If file_path = "" Then
    Err.Raise (448)
End If
'
If (FileType < 1 Or FileType > 9) Then
    Err.Raise (447)
End If
'
If IsTDS5054 And FileName = "" Then
    Err.Raise (450)
End If
'
'set mouse to hourglass when making measurement
Screen.MousePointer = vbHourglass
'
If JustGrat Then
    ScreenShow = "GRAT"
Else
    ScreenShow = "FULLNO"
End If
'
' Default coded for the TDS3000 Series. Any needed changes are when there is a If IsTDS5054, or a MSO4104 (not tested as of version 1.0.0 since it is out for calibration ....
' I think the MSO4104 should work the same as the TDS5054, but that will wait to be seen.  When tested I will updat this Module as-needed.
'
Select Case FileType
    '
    ' FileSize must be greater than the expected size of the resulting file.  These are estimated.  You may have to adjust.  The smaller this value, the fast the routine works.
    ' Too small and the resulting graphics file has clipped black sections.
    '
    ' Ext is the desired extension of the file that will be saved.
    ' FileFormat is to tell the scope what type of file to create.
    '
    Case 1
        Ext = ".PNG"
        FileFormat = "PNG"
        FileSize = 20000#
    Case 2
        Ext = ".BMP"
        If IsTDS5054 Then
            FileFormat = "BMP"
            FileSize = 1000000#
            Palette = "COLO"
        Else
            ' TDS3000
            FileFormat = "BMPCOLOR"
            FileSize = 350000#
        End If
    Case 3
        Ext = ".BMP"
        FileFormat = "BMP"
        If IsTDS5054 Then
            FileSize = 1000000#
            Palette = "BLACKANDWH"
        Else
            ' TDS3000
            FileSize = 350000#
        End If
    Case 4
        Ext = ".TIF"
        FileFormat = "TIFF"
        FileSize = 200000#
    Case 5
        Ext = ".PCX"
        If TDS5054 Then
            FileFormat = "PCX"
        Else
            ' TDS3000
            FileFormat = "PCXCOLOR"
        End If
        FileSize = 150000#
    Case 6
        Ext = ".PCX"
        FileFormat = "PCX"
        FileSize = 150000#
        If TDS5054 Then
            Palette = "BLACKANDWH"
        End If
    Case 7  ' - TDS3000 Series ONLY
        Ext = ".RLE"
        FileFormat = "RLE"
        FileSize = 350000#
    Case 8  ' - TDS5054 ONLY
        Ext = ".JPG"
        FileFormat = "JPEG"
        FileSize = 1000000#
        Palette = "COLOR"
    Case 9  ' - TDS5054 ONLY
        Ext = ".JPG"
        FileFormat = "JPEG"
        FileSize = 1000000#
        Palette = "BLACKANDWH"
End Select
'
If InkSaver Then
    ' TDS3000
    SaveInk = "ON"
    ' TDS5054
    TDS5000SaveInk = "INKSaver"
Else
    ' TDS3000
    SaveInk = "off"
    ' TDS5054
    TDS5000SaveInk = "NORMAL"
End If
'
Set LioMgr = New VisaComLib.ResourceManager
Set equip = New VisaComLib.FormattedIO488
'
Set equip.IO = LioMgr.Open(gpib_address)
'
'set timeout
equip.IO.timeout = 10000 '100 seconds
'
If IsTDS5054 Then
    '
    If FileName = "" Then FileName = "Temp123"
    '
    equip.WriteString "EXPort:FORMat " & FileFormat
    equip.WriteString "EXPort:PALETTE " & Palette
    equip.WriteString "EXPort:IMAG " & TDS5000SaveInk
    equip.WriteString "EXPort:VIEW " & ScreenShow
    equip.WriteString "HARDCOPY:PORT FILE"
    equip.WriteString "HARDCOPY:FILENAME " & Chr(34) & "C:\TekScope\images\" & FileName & Ext & Chr$(34)
    equip.WriteString "HARDCOPY:PORTRAIT"
    equip.WriteString "DIS:COLO:PALETTE USER"
    equip.WriteString "DIS:COLO:PALETTE:USER:CH 180,30,100" 'Have to turn lightness down on TDS5054 to get a decent color display saved to file
    '
Else
    '
    'TDS3000 Series
    '
    equip.WriteString "HARDCOPY:PORT GPIB"
    equip.WriteString "HARDCOPY:FORMAT " & FileFormat
    equip.WriteString "HARDCOPY:INKSAVER " & SaveInk
    equip.WriteString "HARDCOPY:PORTRAIT"
    '
End If
'
'perform hardcopy
equip.WriteString "HARDCOPY START"
If TDS5054 Then
    equip.WriteString "*WAI"
    Do
        '
        equip.WriteString "*OPC?"
        status = equip.ReadString
        '
        DoEvents
        '
    Loop While status <> 1
End If
'
If IsTDS5054 Then
    '
    equip.WriteString "FILESYSTEM:READFILE " & Chr(34) & "C:\TekScope\images\" & FileName & Ext & Chr$(34), True
    '
End If
'
'write file
fn = FreeFile()
Open file_path & Ext For Binary Lock Read Write As #fn
    byteData = equip.IO.Read(FileSize)
    Put #fn, , byteData
Close #fn
'
equip.WriteString "FileSystem:Delete " & Chr(34) & "C:\TekScope\images\" & FileName & Ext & Chr$(34)
'
If IsTDS5054 Then
    equip.WriteString "DIS:COLO:PALETTE:USER:CH 180,50,100"
End If
'
'close equipment
equip.IO.Close
Set equip.IO = Nothing
Set equip = Nothing
'
'set mouse back to default
Screen.MousePointer = vbDefault
'
Exit Function
'
ProcError:
'
If gpib_address = "" Then
    MsgBox "The following error occured: " & Err.Description & ". GPIB Address blank!", vbOKOnly, "!!! ERROR !!!"
    GoTo OutOfHere
End If
'
If file_path = "" Then
    MsgBox "The following error occured: " & Err.Description & ". File_Path can not be blank!", vbOKOnly, "!!! ERROR !!!"
    GoTo OutOfHere
End If
'
If Err = 380 Then
    MsgBox "The following error occured: " & Err.Description & ". GPIB Not properly defined!" & Chr(10) & Chr(10) _
        & "The Address must be specified similar to this: " & Chr$(34) & "GPIB0::1::INSTR" & Chr$(34), vbOKOnly, "!!! ERROR !!!"
    GoTo OutOfHere
End If
'
If Err = 447 Then
    MsgBox "The following error occured: " & Err.Description & ". FileType is NOT within the acceptable range!", vbOKOnly, "!!! ERROR !!!"
    GoTo OutOfHere
End If
'
If Err = 450 Then
    MsgBox "The following error occured: " & Err.Description & ". TDS5054/MSO4104 Specified.  Internal Scope FileName is NOT optional!!!", vbOKOnly, "!!! ERROR !!!"
    GoTo OutOfHere
End If
'
MsgBox "The following error occured: " & Err.Description & " and no furthere description in available.", vbOKOnly, "!!! ERROR !!!"
'
OutOfHere:
    'set mouse back to default
    Screen.MousePointer = vbDefault
    Capture_ScopeScreen = False
    Exit Function
'
End Function
