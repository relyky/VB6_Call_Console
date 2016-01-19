VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'LinearConnections 2002

Private hConsole As Long

Public Function CloseConsole() As Boolean
    If CloseHandle(hConsole) <> 0 Then hConsole = 0
   FreeConsole
  
End Function

Public Function OpenConsole() As Boolean
    
    'We open the Window
    If AllocConsole() Then
        hConsole = GetStdHandle(STD_OUTPUT_HANDLE)
        If hConsole = 0 Then
            FreeConsole
        Else
            'SetConsoleCtrlHandler AddressOf ConsoleHandler, True
            OpenConsole = True
        End If

    End If

End Function

Public Function IsOpen() As Boolean

    IsOpenConsole = IIf(hConsole = 0, False, True)

End Function

Public Function SendText(sText As String) As Boolean
  
    Dim sOut As String, cWritten As Long

    sOut = sText & vbCrLf
    'Write the text to the console
    SendText = WriteConsole(hConsole, ByVal sOut, LenB(StrConv(sOut, vbFromUnicode)), cWritten, ByVal 0&)
 
End Function

Public Sub OpenConsoleAndSendText(sText As String)

    If Not Me.IsOpen Then Me.OpenConsole
    Call Me.SendText(sText)

End Sub

Private Sub Class_Terminate()
    'If a console is open then
    If hConsole <> 0 Then
        'close our handle to it and
        CloseHandle hConsole
        'disconnect it from the process.
        FreeConsole
    End If
End Sub
