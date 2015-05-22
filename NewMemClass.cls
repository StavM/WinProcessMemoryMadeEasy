VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NewMemClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : NewMemClass
' Author    : Stav Mann
' Date      : 17/05/2015
' Mail      : Stavmann2@gmail.com
' Purpose   : Read and Write 32bit process's virtual memory values (ReadProcessMemory&WriteProcessMemory wrapper)
'---------------------------------------------------------------------------------------

Option Explicit

'APIs
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'Premission Constants
Const PROCESS_CREATE_PROCESS As Long = &H80
Const PROCESS_CREATE_THREAD  As Long = &H2
Const PROCESS_VM_OPERATION   As Long = &H8
Const PROCESS_VM_READ        As Long = &H10
Const PROCESS_VM_WRITE       As Long = &H20
Const PROCESS_ALL_ACCESS     As Long = &H1F0FFF

'Special Handles Constants
Const INVALID_HANDLE         As Long = 0

'Inner variables
Private lngPid               As Long        'Holds the process ID
Private lngHandle            As Long        'Holds the process Handle
Private lngOpenHandle_Read   As Long        'Holds the process opened handle - for reading
Private lngOpenHandle_Write  As Long        'Holds the process opened handle - for writing
Private lngThread            As Long        'Holds the thread ID that opened the process

'Type Sizes
Const SIZEOF_BYTE            As Integer = 1
Const SIZEOF_INTEGER         As Integer = 2
Const SIZEOF_LONG            As Integer = 4
Const SIZEOF_DOUBLE          As Integer = 8

'=== Get Process ID ===
Public Property Get ProcessID() As Long
    ProcessID = lngPid
End Property

'=== Get Process Handle ===
Public Property Get ProcessHandle() As Long
    ProcessHandle = lngHandle
End Property

'=== Get Thread Handle ===
Public Property Get ThreadHandle() As Long
    ThreadHandle = lngThread
End Property

'=== Get Open Process's Handle (Reading) ===
Public Property Get OpenProcessHandle_Read() As Long
    OpenProcessHandle_Read = lngOpenHandle_Read
End Property

'=== Get Open Process's Handle (Writing) ===
Public Property Get OpenProcessHandle_Write() As Long
    OpenProcessHandle_Write = lngOpenHandle_Write
End Property

'---------------------------------------------------------------------------------------
' Procedure : AttachToProcess_ByWindowName
' Purpose   : use to specify the exact window name that you want to target for memory reading\writing.
'---------------------------------------------------------------------------------------
Public Sub AttachToProcess_ByWindowName(sExactWindowName As String)
    lngHandle = FindWindow(vbNullString, sExactWindowName)
    lngThread = GetWindowThreadProcessId(lngHandle, lngPid)
    
    'Get two reference handles - one for reading, one for writing
    Call OpenProcessByHandle(lngOpenHandle_Read, PROCESS_VM_READ)
    Call OpenProcessByHandle(lngOpenHandle_Write, PROCESS_VM_WRITE)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AttachToProcess_ByWindowClass
' Purpose   : use to specify the exact class name that you want to target for memory reading\writing.
'---------------------------------------------------------------------------------------
Public Sub AttachToProcess_ByWindowClass(sExactClassName As String)
    lngHandle = FindWindow(sExactClassName, vbNullString)
    lngThread = GetWindowThreadProcessId(lngHandle, lngPid)
    
    'Get two reference handles - one for reading, one for writing
    Call OpenProcessByHandle(lngOpenHandle_Read, PROCESS_VM_READ)
    Call OpenProcessByHandle(lngOpenHandle_Write, PROCESS_VM_WRITE)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AttachToProcess_TopMostWindow
' Purpose   : use to target the current top most window (in foreground) for memory reading\writing, regardless of its name.
'---------------------------------------------------------------------------------------
Public Sub AttachToProcess_TopMostWindow()
    lngHandle = FindWindow(vbNull, vbNull)
    lngThread = GetWindowThreadProcessId(lngHandle, lngPid)
    
    'Get two reference handles - one for reading, one for writing
    Call OpenProcessByHandle(lngOpenHandle_Read, PROCESS_VM_READ)
    Call OpenProcessByHandle(lngOpenHandle_Write, PROCESS_VM_WRITE)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ReleaseProcess
' Purpose   : use to close read, and write, opened handles
'---------------------------------------------------------------------------------------
Public Sub ReleaseProcess()
    Call CloseProcessByHandle(lngOpenHandle_Read)
    Call CloseProcessByHandle(lngOpenHandle_Write)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Read_Byte
' Purpose   : Reads a 1 byte address (byte)
'---------------------------------------------------------------------------------------
Public Function Read_Byte(lngAddressToReadfrom As Long) As Single
    Call ReadMemory(VarPtr(Read_Byte), SIZEOF_BYTE, lngAddressToReadfrom)
End Function

'---------------------------------------------------------------------------------------
' Procedure : Read_Integer
' Purpose   : Reads a 2 byte address (Int)
'---------------------------------------------------------------------------------------
Public Function Read_Integer(lngAddressToReadfrom As Long) As Integer
    Call ReadMemory(VarPtr(Read_Integer), SIZEOF_INTEGER, lngAddressToReadfrom)
End Function

'---------------------------------------------------------------------------------------
' Procedure : Read_Long
' Purpose   : Reads a 4 byte address (Long)
'---------------------------------------------------------------------------------------
Public Function Read_Long(lngAddressToReadfrom As Long) As Long
    Call ReadMemory(VarPtr(Read_Long), SIZEOF_LONG, lngAddressToReadfrom)
End Function

'---------------------------------------------------------------------------------------
' Procedure : Read_Float
' Purpose   : Reads a 4 byte address (Float)
'---------------------------------------------------------------------------------------
Public Function Read_Float(lngAddressToReadfrom As Long) As Single
    Call ReadMemory(VarPtr(Read_Float), SIZEOF_LONG, lngAddressToReadfrom)
End Function

'---------------------------------------------------------------------------------------
' Procedure : Read_Double
' Purpose   : Reads a 8 byte address (double)
'---------------------------------------------------------------------------------------
Public Function Read_Double(lngAddressToReadfrom As Long) As Double
    Call ReadMemory(VarPtr(Read_Double), SIZEOF_DOUBLE, lngAddressToReadfrom)
End Function

'---------------------------------------------------------------------------------------
' Procedure : Read_String
' Purpose   : Reads a 255 long string unless limited by the 2nd argument
'---------------------------------------------------------------------------------------
Public Function Read_String(lngAddressToReadfrom As Long, Optional lngLengthOfString As Long = 255) As String
    Dim strBuffer       As String * 255
    Dim intNullLocation As Long

    strBuffer = String(lngLengthOfString, 0)
    
    'if successfull
    If (ReadProcessMemory(lngOpenHandle_Read, ByVal lngAddressToReadfrom, ByVal strBuffer, lngLengthOfString, &O0) <> 0) Then
    
        'return all chars until the \null char
        intNullLocation = InStr(strBuffer, Chr(0))
        If intNullLocation > 0 Then strBuffer = Left(strBuffer, intNullLocation - 1)
        
        'cleanup
        Read_String = Trim(strBuffer)
    End If
End Function

'Open process with premission level
Private Function OpenProcessByHandle(ByRef lngOpenHandle As Long, Optional lngPremissionLevel As Long = PROCESS_ALL_ACCESS) As Boolean
    'if there's an open handle close it prior to getting a new one.
    If lngOpenHandle <> INVALID_HANDLE Then CloseProcessByHandle (lngOpenHandle)
    'open the process
    lngOpenHandle = OpenProcess(lngPremissionLevel, False, lngPid)
    'validate
    OpenProcessByHandle = (lngOpenHandle <> INVALID_HANDLE)
End Function

'Close process by opened-handle
Private Function CloseProcessByHandle(lngProcessOpenHandle As Long) As Boolean
    CloseProcessByHandle = (CloseHandle(lngProcessOpenHandle) <> INVALID_HANDLE)
End Function

'Read Memory to buffer
Private Function ReadMemory(lngDestinationPointer As Long, lngLength As Long, lngAddress As Long) As Boolean
    ReadMemory = (ReadProcessMemory(lngOpenHandle_Read, ByVal lngAddress, ByVal lngDestinationPointer, lngLength, &O0) <> INVALID_HANDLE)
End Function




