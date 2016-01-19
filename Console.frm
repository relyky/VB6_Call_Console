VERSION 5.00
Begin VB.Form frmConsole 
   BackColor       =   &H80000017&
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "VB6 Console Application by Linear Connections"
   ClientHeight    =   3135
   ClientLeft      =   1095
   ClientTop       =   1500
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   3135
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cbConsoleOutput2 
      Caption         =   "Open Console and Send Text"
      Height          =   465
      Left            =   2100
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   750
      Width           =   3810
   End
   Begin VB.CommandButton cbShellBatch 
      Caption         =   "Shell Console.bat"
      Height          =   465
      Left            =   4095
      TabIndex        =   4
      Top             =   105
      Width           =   1800
   End
   Begin VB.CommandButton cbOpenConsole 
      Caption         =   "Open Console"
      Height          =   465
      Left            =   135
      TabIndex        =   3
      Top             =   105
      Width           =   1800
   End
   Begin VB.CommandButton cbCloseConsole 
      Caption         =   "Close Console"
      Height          =   465
      Left            =   2115
      TabIndex        =   2
      Top             =   105
      Width           =   1800
   End
   Begin VB.TextBox tbConsole 
      Height          =   405
      Left            =   1080
      TabIndex        =   1
      Text            =   "LinearConnections"
      Top             =   1350
      Width           =   5490
   End
   Begin VB.CommandButton cbConsoleOutput 
      Caption         =   "Send Text to Console"
      Height          =   465
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   120
      Width           =   1800
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "Modify by Rely 2016"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   1770
      TabIndex        =   7
      Top             =   2460
      Width           =   4335
   End
   Begin VB.Label lbLen 
      BorderStyle     =   1  '單線固定
      Caption         =   "Len"
      Height          =   345
      Left            =   6720
      TabIndex        =   6
      Top             =   1380
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "Made by LinearConnections 2002"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   1830
      TabIndex        =   5
      Top             =   2070
      Width           =   4335
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Made by LinearConnections 2002
Option Explicit

Public objConsole As New clsConsole

Private Sub cbCloseConsole_Click()

    'Close the console
    objConsole.CloseConsole

End Sub

Private Sub cbConsoleOutput_Click()
'Displays the text...
    objConsole.SendText (tbConsole.Text)
    
End Sub

Private Sub cbConsoleOutput2_Click()

    objConsole.OpenConsoleAndSendText tbConsole.Text

End Sub

Private Sub cbOpenConsole_Click()

    'If we don't successfully open a new console then
    If Not objConsole.OpenConsole Then
        'Send an error msg with an msg box
        MsgBox "Couldn't Start console"
    End If

End Sub

Private Sub cbShellBatch_Click()
    
    'If it's not open yet then,
    If objConsole.IsOpen Then
        'we need to tell the user to open one,
        MsgBox "Please open a console window before running the batch file."
    Else
        'The batch file will help...
        'go to the console.
        Shell """" & App.Path & "\console.bat""", vbNormalFocus
    End If
    
End Sub

Private Sub Form_Load()
    ' This is the default textbox message.
    tbConsole.Text = "LinearConnections - 可以秀中文哦"
    ' show the rght length of the textbox message.
    lbLen.Caption = LenB(StrConv(tbConsole.Text, vbFromUnicode))
  
End Sub

Private Sub tbConsole_Change()

    ' 計算正確的文字長度
    lbLen.Caption = LenB(StrConv(tbConsole.Text, vbFromUnicode))

End Sub
