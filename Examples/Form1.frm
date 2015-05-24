VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1 - Reading Memory"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'don't forget to add reference to the DLL off the \bin folder -
'on the above menus > Project > References... > Browse > go to the bin folder and select 'WinMemoryMadeEasy.dll' > click ok

Option Explicit

Private Sub Form_Load()
    
    'example of a number that will be read
    Const temp As Integer = 1234
    
    'a process's window name
    Dim EXE_WINDOW_NAME As String
        EXE_WINDOW_NAME = Form1.Caption
    
    'an address to read (in this example, the address of the above 'temp' integer)
    Dim ADDRESS_TO_READ As Long
        ADDRESS_TO_READ = VarPtr(temp)
    
    'create object
    Dim memProcess As NewMemClass
    Set memProcess = New NewMemClass
    
    'attach to process
    memProcess.AttachToProcess_ByWindowName (EXE_WINDOW_NAME)
    
    'display messagebox with the content of the memory value within the address space of notepad
    MsgBox memProcess.Read_Integer(ADDRESS_TO_READ)
    
    'cleanup
    memProcess.ReleaseProcess
    Set memProcess = Nothing
    
End Sub
