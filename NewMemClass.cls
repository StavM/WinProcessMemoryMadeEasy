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

'Private variables
Private lngPid        As Long
Private lngHandle     As Long
Private lngOpenHandle As Long
Private lngThread     As Long

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

'=== Get Open Process's Handle ===
Public Property Get OpenProcessHandle() As Long
    OpenProcessHandle = lngOpenHandle
End Property

'---------------------------------------------------------------------------------------
' Procedure : AttachToProcess_ByWindowName
' Purpose   : use to specify the exact window name that you want to target for memory reading\writing.
'---------------------------------------------------------------------------------------
Public Sub AttachToProcess_ByWindowName(sExactWindowName As String)
        lngHandle = FindWindow(vbNullString, sExactWindowName)
        lngThread = GetWindowThreadProcessId(lngHandle, lngPid)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AttachToProcess_ByWindowClass
' Purpose   : use to specify the exact class name that you want to target for memory reading\writing.
'---------------------------------------------------------------------------------------
Public Sub AttachToProcess_ByWindowClass(sExactClassName As String)
        lngHandle = FindWindow(sExactClassName, vbNullString)
        lngThread = GetWindowThreadProcessId(lngHandle, lngPid)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AttachToProcess_TopMostWindow
' Purpose   : use to target the current top most window (in foreground) for memory reading\writing, regardless of its name.
'---------------------------------------------------------------------------------------
Public Sub AttachToProcess_TopMostWindow()
        lngHandle = FindWindow(vbNull, vbNull)
        lngThread = GetWindowThreadProcessId(lngHandle, lngPid)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : OpenWithPremissions()
' Purpose   : used to open a process (by that getting an OpenProcessHandle (lngOpenHandle) to which we can address our APIs when we read\write)
'---------------------------------------------------------------------------------------
Private Sub OpenWithPremissions(Optional lngPremissionType As Long = PROCESS_ALL_ACCESS)
    lngOpenHandle = OpenProcess(lngPremissionType, False, lngPid)
End Sub


