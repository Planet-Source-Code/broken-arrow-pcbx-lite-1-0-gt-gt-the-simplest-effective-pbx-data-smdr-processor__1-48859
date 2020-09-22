Attribute VB_Name = "modMain"
Option Explicit

'Type of the PBX data format variable
Public Type PBXDataMap
    StartMarker1Start As Long
    StartMarker1 As String
    StartMarker2Start As Long
    StartMarker2 As String
    DateStart As Long
    DateLength As Long
    SystemDate As Boolean
    TimeStart As Long
    TimeLength As Long
    SystemTime As Boolean
    ExtensionStart As Long
    ExtensionLength As Long
    COLineStart As Long
    COLineLength As Long
    ForwardingFlagStart As Long
    ForwardingFlag As String
    DialedNumberStart As Long
    DialedNumberLength As Long
    IncomingMarkerStart As Long
    IncomingMarker As String
    CallerStart As Long
    CallerLength As Long
    CallerDontRead As Boolean
    RingDurationStart As Long
    RingDurationLength As Long
    DurationStart As Long
    DurationLength As Long
    AccountStart As Long
    AccountLength As Long
    AccountDontRead As Boolean
    EndMarkerStart As Long
    EndMarker As String
End Type
Public DataMap As PBXDataMap

Private Type COMSetting
    Port As Integer
    BaudRate As Long
    DataBit As Integer
    StopBit As Single
    Parity As String
    HandShaking As Integer
End Type
Public COM As COMSetting

'Dim AppINI As New clsINI
Public COMData As String

Sub Main()
'Check if the application is already running
If App.PrevInstance = True Then
'    MsgBox "PCBX Lite is already running on this system." & vbCrLf & vbCrLf & "Multiple instances are not supported in this version. We are sorry for the inconvenience", vbInformation + vbOKOnly + vbMsgBoxSetForeground, "PCBX Lite 1.0"
'    End
End If

LoadAppSetting 'Load application settings

'Check if the database exists!
If Dir(MDBDatabase) <= "" Then
    MsgBox "The database either missing or moved, please located the database manually. PCBX Lite cannot load without the valid database shipped with it. If this seems an accident, please reinstall the application to reset the database.", vbCritical + vbOKOnly, "Database missing!"
    frmSettings.Show vbModal
'    MsgBox "PCBX Lite will now close. Please rerun PCBX Lite with the new database in effect.", vbInformation + vbOKOnly, "Reload PCBX LIte"
    Shell CheckPath(App.Path) & App.EXEName & ".exe"
    End
End If

GlobalADOConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & MDBDatabase

LoadDataMap 'Set current data map
LoadCOMSetting 'Set current RS232C Serial COM setting

Load frmMain
frmMain.Show


Dim WinState As Long
WinState = frmMain.WindowState
If frmMain.WindowState = vbNormal Then frmMain.WindowState = vbMaximized Else frmMain.WindowState = vbNormal
frmMain.WindowState = WinState
End Sub

Private Sub LoadDataMap()
DataMap.StartMarker1Start = ReadINIInt("Data mapping", "StartMarker1Start")
DataMap.StartMarker1 = ReadINIStr("Data mapping", "StartMarker1")
DataMap.StartMarker2Start = ReadINIInt("Data mapping", "StartMarker2Start")
DataMap.StartMarker2 = ReadINIStr("Data mapping", "StartMarker2")
DataMap.DateStart = ReadINIInt("Data mapping", "DateStart")
DataMap.DateLength = ReadINIInt("Data mapping", "DateLength")
DataMap.SystemDate = CBool(ReadINIStr("Data mapping", "SystemDate"))
DataMap.TimeStart = ReadINIInt("Data mapping", "TimeStart")
DataMap.TimeLength = ReadINIInt("Data mapping", "TimeLength")
DataMap.SystemTime = CBool(ReadINIStr("Data mapping", "SystemTime"))
DataMap.ExtensionStart = ReadINIInt("Data mapping", "ExtensionStart")
DataMap.ExtensionLength = ReadINIInt("Data mapping", "ExtensionLength")
DataMap.COLineStart = ReadINIInt("Data mapping", "COLineStart")
DataMap.COLineLength = ReadINIInt("Data mapping", "COLineLength")
DataMap.ForwardingFlagStart = ReadINIInt("Data mapping", "ForwardingFlagStart")
DataMap.ForwardingFlag = ReadINIStr("Data mapping", "ForwardingFlag")
DataMap.DialedNumberStart = ReadINIInt("Data mapping", "DialedNumberStart")
DataMap.DialedNumberLength = ReadINIInt("Data mapping", "DialedNumberLength")
DataMap.IncomingMarkerStart = ReadINIInt("Data mapping", "IncomingMarkerStart")
DataMap.IncomingMarker = ReadINIStr("Data mapping", "IncomingMarker")
DataMap.CallerStart = ReadINIInt("Data mapping", "CallerStart")
DataMap.CallerLength = ReadINIInt("Data mapping", "CallerLength")
DataMap.CallerDontRead = CBool(ReadINIStr("Data mapping", "CallerDontRead"))
DataMap.RingDurationStart = ReadINIInt("Data mapping", "RingDurationStart")
DataMap.RingDurationLength = ReadINIInt("Data mapping", "RingDurationLength")
DataMap.DurationStart = ReadINIInt("Data mapping", "DurationStart")
DataMap.DurationLength = ReadINIInt("Data mapping", "DurationLength")
DataMap.AccountStart = ReadINIInt("Data mapping", "AccountStart")
DataMap.AccountLength = ReadINIInt("Data mapping", "AccountLength")
DataMap.AccountDontRead = CBool(ReadINIStr("Data mapping", "AccountDontRead"))
DataMap.EndMarkerStart = ReadINIInt("Data mapping", "EndMarkerStart")
DataMap.EndMarker = ReadINIStr("Data mapping", "EndMarker")
End Sub

Public Sub SaveDataMap()
WriteINI "Data mapping", "StartMarker1Start", CStr(DataMap.StartMarker1Start)
WriteINI "Data mapping", "StartMarker1", DataMap.StartMarker1
WriteINI "Data mapping", "StartMarker2Start", CStr(DataMap.StartMarker2Start)
WriteINI "Data mapping", "StartMarker2", DataMap.StartMarker2
WriteINI "Data mapping", "DateStart", CStr(DataMap.DateStart)
WriteINI "Data mapping", "DateLength", CStr(DataMap.DateLength)
WriteINI "Data mapping", "SystemDate", CStr(CInt(DataMap.SystemDate))
WriteINI "Data mapping", "TimeStart", CStr(DataMap.TimeStart)
WriteINI "Data mapping", "TimeLength", CStr(DataMap.TimeLength)
WriteINI "Data mapping", "SystemTime", CStr(CInt(DataMap.SystemTime))
WriteINI "Data mapping", "ExtensionStart", CStr(DataMap.ExtensionStart)
WriteINI "Data mapping", "ExtensionLength", CStr(DataMap.ExtensionLength)
WriteINI "Data mapping", "COLineStart", CStr(DataMap.COLineStart)
WriteINI "Data mapping", "COLineLength", CStr(DataMap.COLineLength)
WriteINI "Data mapping", "ForwardingFlagStart", CStr(DataMap.ForwardingFlagStart)
WriteINI "Data mapping", "ForwardingFlag", DataMap.ForwardingFlag
WriteINI "Data mapping", "DialedNumberStart", CStr(DataMap.DialedNumberStart)
WriteINI "Data mapping", "DialedNumberLength", CStr(DataMap.DialedNumberLength)
WriteINI "Data mapping", "IncomingMarkerStart", CStr(DataMap.IncomingMarkerStart)
WriteINI "Data mapping", "IncomingMarker", DataMap.IncomingMarker
WriteINI "Data mapping", "CallerStart", CStr(DataMap.CallerStart)
WriteINI "Data mapping", "CallerLength", CStr(DataMap.CallerLength)
WriteINI "Data mapping", "CallerDontRead", CStr(CInt(DataMap.CallerDontRead))
WriteINI "Data mapping", "RingDurationStart", CStr(DataMap.RingDurationStart)
WriteINI "Data mapping", "RingDurationLength", CStr(DataMap.RingDurationLength)
WriteINI "Data mapping", "DurationStart", CStr(DataMap.DurationStart)
WriteINI "Data mapping", "DurationLength", CStr(DataMap.DurationLength)
WriteINI "Data mapping", "AccountStart", CStr(DataMap.AccountStart)
WriteINI "Data mapping", "AccountLength", CStr(DataMap.AccountLength)
WriteINI "Data mapping", "AccountDontRead", CStr(CInt(DataMap.AccountDontRead))
WriteINI "Data mapping", "EndMarkerStart", CStr(DataMap.EndMarkerStart)
WriteINI "Data mapping", "EndMarker", DataMap.EndMarker
End Sub

Private Sub LoadCOMSetting()
COM.Port = ReadINIInt("RS232C Serial COM", "Port")
COM.BaudRate = ReadINIInt("RS232C Serial COM", "BaudRate")
COM.DataBit = ReadINIInt("RS232C Serial COM", "DataBit")
COM.StopBit = ReadINIInt("RS232C Serial COM", "StopBit")
COM.Parity = ReadINIStr("RS232C Serial COM", "Parity")
COM.HandShaking = ReadINIInt("RS232C Serial COM", "HandShaking")
End Sub

Public Sub SaveCOMSetting()
WriteINI "RS232C Serial COM", "Port", CStr(COM.Port)
WriteINI "RS232C Serial COM", "BaudRate", CStr(COM.BaudRate)
WriteINI "RS232C Serial COM", "DataBit", CStr(COM.DataBit)
WriteINI "RS232C Serial COM", "StopBit", CStr(COM.StopBit)
WriteINI "RS232C Serial COM", "Parity", CStr(COM.Parity)
WriteINI "RS232C Serial COM", "HandShaking", CStr(COM.HandShaking)
End Sub

Public Sub SaveAppSetting()
WriteINI "Application Setting", "Database", MDBDatabase
End Sub

Private Sub LoadAppSetting()
MDBDatabase = ReadINIStr("Application Setting", "Database", , CheckPath(App.Path, True) & "PCBXLite.mdb")
End Sub
