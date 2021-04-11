VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chatty - Main"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   7710
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtchat 
      Height          =   4575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   16
      Top             =   0
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Connection Status"
      Height          =   1215
      Left            =   4800
      TabIndex        =   14
      Top             =   1200
      Width           =   2775
      Begin VB.Label lblconnection 
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SCC"
      Height          =   1215
      Left            =   4800
      TabIndex        =   11
      Top             =   2520
      Width           =   2775
      Begin VB.Label lblscc 
         BackStyle       =   0  'Transparent
         Caption         =   " Disabled"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblsecure 
         Alignment       =   2  'Center
         Caption         =   "Secure Communication Channel"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2535
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   960
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtsay 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Width           =   7335
   End
   Begin VB.OptionButton opttype 
      Caption         =   "Server"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   2
      Top             =   4680
      Width           =   855
   End
   Begin VB.OptionButton opttype 
      Caption         =   "Client"
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   1
      Top             =   4680
      Value           =   -1  'True
      Width           =   855
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   5280
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtip 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmddisconnect 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdlisten 
      Caption         =   "Listen"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdconnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   6000
      X2              =   7440
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   6000
      X2              =   7440
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label1 
      Caption         =   "ChAtty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4920
      TabIndex        =   21
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   735
      Left            =   4800
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label lblreceived 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   6240
      TabIndex        =   20
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label lblsent 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   6000
      TabIndex        =   19
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblbreceive 
      Caption         =   "Bytes Received:"
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label lblbsent 
      Caption         =   "Bytes Sent:"
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lbltitle 
      Height          =   735
      Left            =   5040
      TabIndex        =   10
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lbltext 
      Caption         =   "Text to Send:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label lblip 
      Caption         =   "IP Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   975
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuclear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnufbar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuaction 
      Caption         =   "Action"
      Begin VB.Menu mnuconnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnulisten 
         Caption         =   "Listen"
      End
      Begin VB.Menu mnudisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnufbar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnumessage 
         Caption         =   "Message"
      End
      Begin VB.Menu mnufbar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnusettings 
         Caption         =   "Settings"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnureadme 
         Caption         =   "Readme"
      End
      Begin VB.Menu mnusupport 
         Caption         =   "Support"
      End
      Begin VB.Menu mnufbar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dot As Byte
Dim nm As String
Dim chatter As String
Dim cht As String
Dim sdr As String
Dim rcv As Long

Private Sub cmdconnect_Click()
dot = 0
If txtip.Text = "" Then
  MsgBox "Please enter an IP address", vbCritical, "Error !"
  Exit Sub
End If
For k = 1 To Len(txtip.Text)
  a = Mid(txtip.Text, k, 1)
  If a = "." Then
    dot = dot + 1
  End If
Next
If dot <> 3 Then
  MsgBox "Invalid IP address format", vbCritical, "Error !"
  Exit Sub
End If
ws.RemotePort = 45660
ws.RemoteHost = txtip.Text
ws.Connect
cmdconnect.Enabled = False
mnuconnect.Enabled = False
cmddisconnect.Enabled = True
mnudisconnect.Enabled = True
txtip.Enabled = False
txtip.BackColor = &HC0C0C0
For k = 0 To 1
  opttype(k).Enabled = False
Next
lblconnection.Caption = "Connecting to " & ws.RemoteHost & "...."
End Sub

Private Sub cmddisconnect_Click()
If ws.State <> sckClosed Then
  If ws.State = sckConnected Then
    ws.SendData "BYE"
    sen = Val(lblsent.Caption) + 3
    lblsent.Caption = Str(sen)
  End If
  DoEvents
  ws.Close
  txtchat.Text = txtchat.Text & "Chat ended - " & Date & ", " & Time & vbNewLine
  If logg = "Log=1" Then
      Open App.Path & "\" & Year(Now) & Month(Now) & Day(Now) & ".txt" For Output As #1
      Print #1, txtchat.Text
      Close #1
  End If
  lblconnection.Caption = " Ready"
  cmddisconnect.Enabled = False
  mnudisconnect.Enabled = False
  If opttype(1).Value = True Then
    cmdlisten.Enabled = True
    mnulisten.Enabled = True
    cmdconnect.Enabled = False
    mnuconnect.Enabled = False
    txtip.Enabled = False
    txtip.BackColor = &HC0C0C0
  ElseIf opttype(0).Value = True Then
    cmdlisten.Enabled = False
    mnulisten.Enabled = False
    cmdconnect.Enabled = True
    mnuconnect.Enabled = True
    txtip.Enabled = True
    txtip.BackColor = vbWhite
  End If
  For k = 0 To 1
    opttype(k).Enabled = True
  Next
  txtsay.Enabled = False
  txtsay.BackColor = &HC0C0C0
End If
End Sub

Private Sub cmdexit_Click()
If ws.State = sckConnected Then
  s = MsgBox("You are currently connected. Are you sure you want to quit ?", vbInformation + vbYesNo, "Confirm Exit")
  If s = vbYes Then
    Unload Me
  End If
End If
End Sub

Private Sub cmdlisten_Click()
ws.LocalPort = 45660
ws.Listen
If ws.State = sckListening Then
  cmdlisten.Enabled = False
  mnulisten.Enabled = False
  cmddisconnect.Enabled = True
  mnudisconnect.Enabled = True
  For k = 0 To 1
    opttype(k).Enabled = False
  Next
  lblconnection.Caption = "Listening on port " & ws.LocalPort & "...."
End If
End Sub

Private Sub cmdmessage_Click()
frmmessage.Show
End Sub

Private Sub Form_Load()
mnulisten.Enabled = False
mnudisconnect.Enabled = False
Open "name.cfg" For Input As #1
Line Input #1, nm
Close #1
Me.Caption = "Chatty - " & nm
lbltitle.Caption = "Chatty v1.10" & vbNewLine
lbltitle.Caption = lbltitle.Caption & vbNewLine & "Local IP Address: " & ws.LocalIP
lblconnection.Caption = " Ready"
txtsay.Enabled = False
txtsay.BackColor = &HC0C0C0
key(1) = 35429567
key(2) = 21444671
key(3) = 31393357
p = 3613
q = 8689
PHI = 31381056
If Dir(App.Path & "\chatty.cfg") = "" Then
  MsgBox "Configuration file not found.", vbCritical, "Error !"
  Exit Sub
End If
Open App.Path & "\chatty.cfg" For Input As #2
Line Input #2, sc
Line Input #2, logg
Close #2
If sc = "SCC=1" Then
  frmmain.Shape1.FillColor = &H8000&
  frmmain.lblscc.Caption = "   Active"
Else
  frmmain.Shape1.FillColor = &H808080
  frmmain.lblscc.Caption = " Disabled"
End If
End Sub

Private Sub mnuabout_Click()
frmabout.Show vbModal
End Sub

Private Sub mnuclear_Click()
txtchat.Text = ""
If logg = "Log=1" Then
  Open App.Path & "\" & Year(Now) & Month(Now) & Day(Now) & ".txt" For Output As #1
  Print #1, ""
  Close #1
End If
End Sub

Private Sub mnuconnect_Click()
Call cmdconnect_Click
End Sub

Private Sub mnudisconnect_Click()
Call cmddisconnect_Click
End Sub

Private Sub mnuexit_Click()
Call cmdexit_Click
End Sub

Private Sub mnulisten_Click()
Call cmdlisten_Click
End Sub

Private Sub mnumessage_Click()
frmmessage.Show
End Sub

Private Sub mnureadme_Click()
If Dir(App.Path & "\Readme.txt") = "" Then
  MsgBox "Readme file not found.", vbCritical, "Error !"
  Exit Sub
End If
Call Shell("notepad.exe Readme.txt", vbNormalFocus)
End Sub

Private Sub mnusettings_Click()
frmsettings.Show
End Sub

Private Sub mnusupport_Click()
frmsupport.Show
End Sub

Private Sub opttype_Click(Index As Integer)
Select Case Index
Case 0:
  txtip.Enabled = True
  txtip.BackColor = vbWhite
  cmdlisten.Enabled = False
  mnulisten.Enabled = False
  cmdconnect.Enabled = True
  mnuconnect.Enabled = True
Case 1:
  txtip.Enabled = False
  txtip.BackColor = &HC0C0C0
  cmdconnect.Enabled = False
  mnuconnect.Enabled = False
  cmdlisten.Enabled = True
  mnulisten.Enabled = True
End Select
End Sub

Private Sub txtsay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If ws.State = sckConnected Then
    If frmmain.lblscc.Caption = "   Active" Then
      sdr = enc(txtsay.Text, key(1), key(3))
      ws.SendData caesarE(sdr) & "TXT-S"
      sen = Val(lblsent.Caption) + Len(caesarE(sdr)) + 5
      lblsent.Caption = Str(sen)
    Else
      ws.SendData txtsay.Text & "TXT"
      sen = Val(lblsent.Caption) + Len(txtsay.Text) + 3
      lblsent.Caption = Str(sen)
    End If
    txtchat.Text = txtchat.Text & nm & ": " & txtsay.Text & vbNewLine
    If logg = "Log=1" Then
      Open App.Path & "\" & Year(Now) & Month(Now) & Day(Now) & ".txt" For Output As #1
      Print #1, txtchat.Text
      Close #1
    End If
    txtsay.Text = ""
  Else
    MsgBox "No connection. Unable to send message", vbCritical, "Error !"
  End If
End If
End Sub

Private Sub ws_Connect()
lblconnection.Caption = "Connected to " & ws.RemoteHost & " on port 45660"
If frmmain.lblscc.Caption = "   Active" Then
  sdr = enc(nm, key(1), key(3))
  ws.SendData caesarE(sdr) & "NM-S"
  sen = Val(lblsent.Caption) + Len(caesarE(sdr)) + 4
  lblsent.Caption = Str(sen)
Else
  ws.SendData nm & "NM"
  sen = Val(lblsent.Caption) + Len(nm) + 2
  lblsent.Caption = Str(sen)
End If
End Sub

Private Sub ws_ConnectionRequest(ByVal requestID As Long)
If ws.State <> sckClosed Then
  ws.Close
End If
ws.Accept requestID
lblconnection.Caption = "Connected to " & ws.RemoteHostIP
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
DoEvents
ws.GetData dat, vbString
rcv = Val(lblreceived.Caption) + bytesTotal
lblreceived.Caption = Str(rcv)

If dat = "BYE" Then
  txtchat.Text = txtchat.Text & "Chat ended - " & Date & ", " & Time & vbNewLine
  If logg = "Log=1" Then
      Open App.Path & "\" & Year(Now) & Month(Now) & Day(Now) & ".txt" For Output As #1
      Print #1, txtchat.Text
      Close #1
  End If
  If ws.State <> sckClosed Then
    ws.Close
  End If
  lblconnection.Caption = " Ready"
  cmddisconnect.Enabled = False
  If opttype(1).Value = True Then
    cmdlisten.Enabled = True
    cmdconnect.Enabled = False
    txtip.Enabled = False
    txtip.BackColor = &HC0C0C0
  ElseIf opttype(0).Value = True Then
    cmdlisten.Enabled = False
    cmdconnect.Enabled = True
    txtip.Enabled = True
    txtip.BackColor = vbWhite
  End If
  For k = 0 To 1
    opttype(k).Enabled = True
  Next
  txtsay.Enabled = False
  txtsay.BackColor = &HC0C0C0
End If

If (Right(dat, 5) = "MSG-S") Or (Right(dat, 3) = "MSG") Then
  If Right(dat, 1) = "S" Then
    recv1 = caesarD(Mid(dat, 1, Len(dat) - 5))
    recv = dec(recv1, key(2), key(3))
  Else
    recv = Mid(dat, 1, Len(dat) - 3)
  End If
  MsgBox recv, , "Message from " & chatter
End If

If (Right(dat, 4) = "RP-S") Or (Right(dat, 2) = "RP") Then
  If Right(dat, 1) = "S" Then
    recv1 = caesarD(Mid(dat, 1, Len(dat) - 4))
    chatter = dec(recv1, key(2), key(3))
  Else
    chatter = Mid(dat, 1, Len(dat) - 2)
  End If
  txtchat.Text = ""
  txtchat.Text = txtchat.Text & "Chat started - " & Date & ", " & Time & vbNewLine
  If logg = "Log=1" Then
      Open App.Path & "\" & Year(Now) & Month(Now) & Day(Now) & ".txt" For Output As #1
      Print #1, txtchat.Text
      Close #1
  End If
  txtsay.Enabled = True
  txtsay.BackColor = vbWhite
End If

If (Right(dat, 4) = "NM-S") Or (Right(dat, 2) = "NM") Then
  If Right(dat, 1) = "S" Then
    recv1 = caesarD(Mid(dat, 1, Len(dat) - 4))
    chatter = dec(recv1, key(2), key(3))
    sdr = enc(nm, key(1), key(3))
    ws.SendData caesarE(sdr) & "RP-S"
    sen = Val(lblsent.Caption) + Len(caesarE(sdr)) + 4
    lblsent.Caption = Str(sen)
  Else
    chatter = Mid(dat, 1, Len(dat) - 2)
    ws.SendData nm & "RP"
    sen = Val(lblsent.Caption) + Len(nm) + 2
    lblsent.Caption = Str(sen)
  End If
  txtchat.Text = ""
  txtchat.Text = txtchat.Text & "Chat started - " & Date & ", " & Time & vbNewLine
  If logg = "Log=1" Then
      Open App.Path & "\" & Year(Now) & Month(Now) & Day(Now) & ".txt" For Output As #1
      Print #1, txtchat.Text
      Close #1
  End If
  txtsay.Enabled = True
  txtsay.BackColor = vbWhite
End If

If (Right(dat, 5) = "TXT-S") Or (Right(dat, 3) = "TXT") Then
  If Right(dat, 1) = "S" Then
    recv1 = caesarD(Mid(dat, 1, Len(dat) - 5))
    cht = dec(recv1, key(2), key(3))
    txtchat.Text = txtchat.Text & chatter & ": " & cht & vbNewLine
    If logg = "Log=1" Then
      Open App.Path & "\" & Year(Now) & Month(Now) & Day(Now) & ".txt" For Output As #1
      Print #1, txtchat.Text
      Close #1
    End If
  Else
    cht = Mid(dat, 1, Len(dat) - 3)
    txtchat.Text = txtchat.Text & chatter & ": " & cht & vbNewLine
    If logg = "Log=1" Then
      Open App.Path & "\" & Year(Now) & Month(Now) & Day(Now) & ".txt" For Output As #1
      Print #1, txtchat.Text
      Close #1
    End If
  End If
End If
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Description, vbCritical, "Error !"
If ws.State <> sckClosed Then
  ws.Close
End If
  lblconnection.Caption = " Ready"
  cmddisconnect.Enabled = False
  mnudisconnect.Enabled = False
  If opttype(1).Value = True Then
    cmdlisten.Enabled = True
    mnulisten.Enabled = True
    cmdconnect.Enabled = False
    mnudisconnect.Enabled = False
    txtip.Enabled = False
    txtip.BackColor = &HC0C0C0
  ElseIf opttype(0).Value = True Then
    cmdlisten.Enabled = False
    mnulisten.Enabled = False
    cmdconnect.Enabled = True
    mnuconnect.Enabled = True
    txtip.Enabled = True
    txtip.BackColor = vbWhite
  End If
  For k = 0 To 1
    opttype(k).Enabled = True
  Next
End Sub

