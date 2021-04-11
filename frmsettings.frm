VERSION 5.00
Begin VB.Form frmsettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   Icon            =   "frmsettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5205
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Logging"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4935
      Begin VB.CheckBox chklog 
         Caption         =   "Logging enabled"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Provides automatic logging of chat. Saves contents of chat to text file"
         Height          =   495
         Left            =   1080
         TabIndex        =   5
         Top             =   840
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SCC"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.CheckBox chksecure 
         Caption         =   "Secure Communication"
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   $"frmsettings.frx":0442
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmsettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chksecure_Click()
If chksecure.Value = Checked Then
  frmmain.Shape1.FillColor = &H8000&
  frmmain.lblscc.Caption = "   Active"
ElseIf chksecure.Value = Unchecked Then
  frmmain.Shape1.FillColor = &H808080
  frmmain.lblscc.Caption = " Disabled"
End If
End Sub

Private Sub cmdok_Click()
Open App.Path & "\chatty.cfg" For Output As #1
If chksecure.Value = Checked Then
  Print #1, "SCC=1"
Else
  Print #1, "SCC=0"
End If
If chklog.Value = Checked Then
  Print #1, "Log=1"
Else
  Print #1, "Log=0"
End If
Close #1
MsgBox "Settings changed. Please restart Chatty for the new settings to take effect", vbInformation, "Change Settings"
Unload Me
End Sub

Private Sub Form_Load()
If Dir(App.Path & "\chatty.cfg") = "" Then
  MsgBox "Configuration file not found.", vbCritical, "Error !"
  Exit Sub
End If
Open App.Path & "\chatty.cfg" For Input As #2
Line Input #2, sc
Line Input #2, lg
Close #2
If sc = "SCC=1" Then
  chksecure.Value = Checked
Else
  chksecure.Value = Unchecked
End If
If lg = "Log=1" Then
  chklog.Value = Checked
Else
  chklog.Value = Unchecked
End If
End Sub
