Attribute VB_Name = "Module_exec"
Public Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Public Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadID As Long
End Type

Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
   hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Declare Function CreateProcessA Lib "kernel32" (ByVal _
   lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
   lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, lpProcessInformation As _
   PROCESS_INFORMATION) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal _
   hObject As Long) As Long

Public Const NORMAL_PRIORITY_CLASS As Long = &H20
Public Const STARTF_USESHOWWINDOW As Long = &H1
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_HIDE As Integer = 0&
Public Const INFINITE = -1&


' *******************  ExecAndWait till exec is terminated  **********************
Public Sub ExecAndWait(cmdline$)
   Dim proc As PROCESS_INFORMATION
   Dim start As STARTUPINFO
   Dim ret As Long

   ' Initialize the STARTUPINFO structure:
   start.cb = Len(start)
   start.dwFlags = STARTF_USESHOWWINDOW
   'start.wShowWindow = SW_SHOWMAXIMIZED
   start.wShowWindow = SW_HIDE
   start.lpTitle = "START"

   ' Start the shelled application:
   ret& = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
      NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

   ' Wait for the shelled application to finish:
   ret& = WaitForSingleObject(proc.hProcess, INFINITE)
   ret& = CloseHandle(proc.hThread)
   ret& = CloseHandle(proc.hProcess)
End Sub


