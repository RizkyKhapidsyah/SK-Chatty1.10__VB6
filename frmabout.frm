VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Chatty"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "frmabout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   3750
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtinfo 
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label lblversion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1425
      Left            =   0
      Picture         =   "frmabout.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdok_Click()
Unload Me
End Sub

Private Sub Form_Load()
With txtinfo
.Text = txtinfo.Text & "Designed and Developed by Benny T." & vbNewLine
.Text = txtinfo.Text & "Modified by Rizky Khapidsyah" & vbNewLine
.Text = txtinfo.Text & "Data Transmission: TCP/IP - Winsock" & vbNewLine
.Text = txtinfo.Text & "Port: 45660 TCP" & vbNewLine
.Text = txtinfo.Text & "Uses Secure Communication Channel (SCC)" & vbNewLine
End With
End Sub


