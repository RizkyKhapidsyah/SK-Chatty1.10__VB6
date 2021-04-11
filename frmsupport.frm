VERSION 5.00
Begin VB.Form frmsupport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Support"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   Icon            =   "frmsupport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4290
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblfree 
      Caption         =   "This program is distributed as freeware"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label lblemail 
      Caption         =   "allegro16@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      MouseIcon       =   "frmsupport.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "E-mail me at :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblinfo 
      Caption         =   $"frmsupport.frx":074C
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmsupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdok_Click()
Unload Me
End Sub

Private Sub lblemail_Click()
d = ShellExecute(0, vbNullString, "mailto:allegro16@hotmail.com", vbNullString, vbNullString, vbNormalFocus)
End Sub

