VERSION 5.00
Begin VB.Form frmmessage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instant Messaging"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   Icon            =   "frmmessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   3960
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdsend 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtmsg 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label lblmessage 
      Caption         =   "Please enter your message:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmmessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim re As String

Private Sub cmdsend_Click()
On Error Resume Next
If txtmsg.Text = "" Then
  MsgBox "Please enter your message", vbCritical, "Error !"
  Exit Sub
End If
If frmmain.ws.State = sckConnected Then
  If frmmain.lblscc.Caption = "   Active" Then
    re = enc(txtmsg.Text, key(1), key(3))
    frmmain.ws.SendData caesarE(re) & "MSG-S"
    sen = Val(frmmain.lblsent.Caption) + Len(caesarE(re)) + 5
    frmmain.lblsent.Caption = Str(sen)
  Else
    frmmain.ws.SendData txtmsg.Text & "MSG"
    sen = Val(frmmain.lblsent.Caption) + Len(txtmsg.Text) + 3
    frmmain.lblsent.Caption = Str(sen)
  End If
  Unload Me
Else
  MsgBox "No connection. Unable to send message", vbCritical, "Error !"
End If
End Sub
