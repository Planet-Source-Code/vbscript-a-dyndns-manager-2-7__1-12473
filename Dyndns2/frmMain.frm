VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DynDNS Manager"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   18
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtPlaceholder 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   "frmMain.frx":0442
      Top             =   5640
      Width           =   5775
   End
   Begin VB.CommandButton cmdClearForm 
      Caption         =   "Clear Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Frame fraAccounts 
      Caption         =   "DynDNS Hosts"
      Height          =   855
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   3975
      Begin VB.ListBox lstHosts 
         Height          =   450
         ItemData        =   "frmMain.frx":0473
         Left            =   120
         List            =   "frmMain.frx":0475
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdRefAccounts 
      Caption         =   "Refresh Accounts"
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   19
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelInfo 
      Caption         =   "Delete Information"
      Height          =   375
      Left            =   4200
      TabIndex        =   17
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpInfo 
      Caption         =   "Update Information"
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSaveInfo 
      Caption         =   "Save Information"
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdRefreshIP 
      Caption         =   "Refresh IP Address"
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update DynDNS"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   240
      Width           =   1695
   End
   Begin VB.Frame fraBackMX 
      Caption         =   "Backup MX"
      Height          =   1095
      Left            =   2160
      TabIndex        =   29
      Top             =   4440
      Width           =   1935
      Begin VB.OptionButton optBackON 
         Caption         =   "Backup MX ON"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton optBackOFF 
         Caption         =   "Backup MX OFF"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame fraWildcard 
      Caption         =   "Wildcard"
      Height          =   1095
      Left            =   120
      TabIndex        =   28
      Top             =   4440
      Width           =   1935
      Begin VB.OptionButton optWildOFF 
         Caption         =   "Wildcard OFF"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optWildOn 
         Caption         =   "Wildcard ON"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame fraDynDNS 
      Caption         =   "DynDNS Information"
      Height          =   1935
      Left            =   120
      TabIndex        =   23
      Top             =   2400
      Width           =   3975
      Begin VB.ComboBox cmbHostType 
         Height          =   315
         ItemData        =   "frmMain.frx":0477
         Left            =   1320
         List            =   "frmMain.frx":0484
         TabIndex        =   6
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtMX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtHost 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblHostType 
         Caption         =   "System"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblMailEx 
         Caption         =   "Mail Exchanger"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblipAddress 
         Caption         =   "IP Address"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblHost 
         Caption         =   "Host Name"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraUserInfo 
      Caption         =   "User Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   3975
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtUsername 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblUsername 
         Caption         =   "Username"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
   End
   Begin MSWinsockLib.Winsock wskDynDNS 
      Left            =   4200
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser wwwFile 
      Height          =   855
      Left            =   120
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5640
      Visible         =   0   'False
      Width           =   5775
      ExtentX         =   10186
      ExtentY         =   1508
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Public SendURL As String
    Dim db As Database
    Dim rs As Recordset

Private Sub Form_Load()
    Set db = OpenDatabase(App.Path & "\dyndns.mdb")
    Set rs = db.OpenRecordset("tblData")
    cmdRefreshIP_Click
    cmdRefAccounts_Click
End Sub

Private Sub cmdTest_Click()
    On Error GoTo ErrorChk
    Dim Username, Password, Hostname, IPAddress As String
    Dim MailExchanger, Wildcard, BackupMX As String
    Dim HostType As String, X As Integer
    Dim Answer As Variant, TheDate As Date, DbDate As Date
    X = 1
    TheDate = Date
    DbDate = rs.Fields("LastRan")
    Username = txtUsername.Text
    Password = txtPassword.Text
    Hostname = txtHost.Text
    IPAddress = txtIP.Text
    MailExchanger = txtMX.Text
    If optWildOFF.Value = True Then
        Wildcard = "OFF"
    ElseIf optWildOn.Value = True Then
        Wildcard = "ON"
    End If
    If optBackOFF.Value = True Then
        BackupMX = "OFF"
    ElseIf optBackON.Value = True Then
        BackupMX = "ON"
    End If
    If cmbHostType.Text = "Dynamic DNS - dyndns" Then
        HostType = "dyndns"
    ElseIf cmbHostType.Text = "Static DNS - statdns" Then
        HostType = "statdns"
    Else
        HostType = "dyndns"
    End If
    If TheDate <= DbDate Then
        Answer = MsgBox("It has been less than 30 days since your last update." & vbCrLf & "Would you still like to perform the update?", vbYesNo + vbCritical, "Continue?")
        Select Case Answer
            Case 6
                SendURL = "http://" & Username & ":" & Password & "@members.dyndns.org/nic/update?system=" & HostType & "&hostname=" & Hostname & "&myip=" & IPAddress & "&wildcard=" & Wildcard & "&mx=" & MailExchanger & "&backmx=" & BackupMX
                Clipboard.Clear
                Clipboard.SetText (SendURL)
                MsgBox SendURL, vbOKOnly, "SendURL Information"
                cmdUpInfo_Click
            Case 7
                MsgBox "Update not performed.", vbOKOnly + vbInformation, "Complete"
            Case Else
                MsgBox "Update not performed.", vbOKOnly + vbInformation, "Complete"
        End Select
    Else
        SendURL = "http://" & Username & ":" & Password & "@members.dyndns.org/nic/update?system=" & HostType & "&hostname=" & Hostname & "&myip=" & IPAddress & "&wildcard=" & Wildcard & "&mx=" & MailExchanger & "&backmx=" & BackupMX
        Clipboard.Clear
        Clipboard.SetText (SendURL)
        MsgBox SendURL, vbOKOnly, "SendURL Information"
        cmdUpInfo_Click
    End If
    Exit Sub
ErrorChk:
    MsgBox "Please fill in at least the following information:" & vbCrLf & "Username, Password, Host Name, IP and Host Type." & vbCrLf & "Or, select an account with the correct information from the list.", vbCritical + vbOKOnly, "Error: Invalid Information"
    Exit Sub
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo ErrorChk
    Dim Username, Password, Hostname, IPAddress As String
    Dim MailExchanger, Wildcard, BackupMX As String
    Dim HostType As String, X As Integer
    Dim Answer As Variant, TheDate As Date, DbDate As Date
    X = 1
    TheDate = Date
    DbDate = rs.Fields("LastRan")
    Username = txtUsername.Text
    Password = txtPassword.Text
    Hostname = txtHost.Text
    IPAddress = txtIP.Text
    MailExchanger = txtMX.Text
    If optWildOFF.Value = True Then
        Wildcard = "OFF"
    ElseIf optWildOn.Value = True Then
        Wildcard = "ON"
    End If
    If optBackOFF.Value = True Then
        BackupMX = "OFF"
    ElseIf optBackON.Value = True Then
        BackupMX = "ON"
    End If
    If cmbHostType.Text = "Dynamic DNS - dyndns" Then
        HostType = "dyndns"
    ElseIf cmbHostType.Text = "Static DNS - statdns" Then
        HostType = "statdns"
    Else
        HostType = "custom"
    End If
    If TheDate <= DbDate Then
        Answer = MsgBox("It has been less than 30 days since your last update." & vbCrLf & "Would you still like to perform the update?", vbYesNo + vbCritical, "Continue?")
        Select Case Answer
            Case 6
                SendURL = "http://" & Username & ":" & Password & "@members.dyndns.org/nic/update?system=" & HostType & "&hostname=" & Hostname & "&myip=" & IPAddress & "&wildcard=" & Wildcard & "&mx=" & MailExchanger & "&backmx=" & BackupMX
                wwwFile.Navigate SendURL
                cmdUpInfo_Click
            Case 7
                MsgBox "Update not performed.", vbOKOnly + vbInformation, "Complete"
            Case Else
                MsgBox "Update not performed.", vbOKOnly + vbInformation, "Complete"
        End Select
    Else
        SendURL = "http://" & Username & ":" & Password & "@members.dyndns.org/nic/update?system=" & HostType & "&hostname=" & Hostname & "&myip=" & IPAddress & "&wildcard=" & Wildcard & "&mx=" & MailExchanger & "&backmx=" & BackupMX
        wwwFile.Navigate SendURL
        cmdUpInfo_Click
    End If
    txtPlaceholder.Visible = False
    wwwFile.Visible = True
    Exit Sub
ErrorChk:
    MsgBox "Please fill in at least the following information:" & vbCrLf & "Username, Password, Host Name, IP and Host Type." & vbCrLf & "Or, select an account with the correct information from the list.", vbCritical + vbOKOnly, "Error: Invalid Information"
    Exit Sub
End Sub

Private Sub cmdRefreshIP_Click()
    txtIP.Text = ""
    txtIP.Text = wskDynDNS.LocalIP
End Sub

Private Sub cmdRefAccounts_Click()
    Dim Item As String
    lstHosts.Clear
    rs.MoveFirst
    Do While Not rs.EOF
        Item = rs.Fields("Host")
        lstHosts.AddItem Item
        rs.MoveNext
    Loop
    txtHost.Text = ""
    txtUsername.Text = ""
    txtPassword.Text = ""
    txtMX.Text = ""
    optWildOFF = True
    optBackOFF = True
    cmdRefreshIP_Click
End Sub

Private Sub cmdClearForm_Click()
    cmdRefreshIP_Click
    cmdRefAccounts_Click
    txtHost.Text = ""
    txtUsername.Text = ""
    txtPassword.Text = ""
    txtMX.Text = ""
    cmbHostType.Text = ""
    optWildOFF = True
    optBackOFF = True
    wwwFile.Visible = False
    txtPlaceholder.Visible = True
End Sub

Private Sub cmdSaveInfo_Click()
    With rs
        .AddNew
            !Host = txtHost.Text
            !Username = Encode(txtUsername.Text)
            !Password = Encode(txtPassword.Text)
            !IPAdd = txtIP.Text
            !LastRan = Date
            If txtMX.Text = "" Then
                !MXRecord = "N/A"
            Else
                !MXRecord = txtMX.Text
            End If
            If cmbHostType.Text = "Dynamic DNS - dyndns" Or cmbHostType.Text = "Static DNS - statdns" Then
                If cmbHostType.Text = "Dynamic DNS - dyndns" Then
                    !HostType = "dyndns"
                ElseIf cmbHostType.Text = "Static DNS - statdns" Then
                    !HostType = "statdns"
                Else
                    !HostType = "custom"
                End If
            Else
                !HostType = "dyndns"
            End If
            If optWildOFF.Value = True Then
                !Wildcard = "OFF"
            ElseIf optWildOn.Value = True Then
                !Wildcard = "ON"
            End If
            If optBackOFF.Value = True Then
                !BackupMX = "OFF"
            ElseIf optBackON.Value = True Then
                !BackupMX = "ON"
            End If
        .Update
    End With
    cmdClearForm_Click
End Sub

Private Sub cmdUpInfo_Click()
    With rs
        .Edit
            !Host = txtHost.Text
            !Username = Encode(txtUsername.Text)
            !IPAdd = txtIP.Text
            !LastRan = Date
            !Password = Encode(txtPassword.Text)
            If cmbHostType.Text = "Dynamic DNS - dyndns" Or cmbHostType.Text = "Static DNS - statdns" Then
                If cmbHostType.Text = "Dynamic DNS - dyndns" Then
                    !HostType = "dyndns"
                ElseIf cmbHostType.Text = "Static DNS - statdns" Then
                    !HostType = "statdns"
                Else
                    !HostType = "custom"
                End If
            Else
                !HostType = "dyndns"
            End If
            If txtMX.Text = "" Then
                !MXRecord = "N/A"
            Else
                !MXRecord = txtMX.Text
            End If
            If optWildOFF.Value = True Then
                !Wildcard = "OFF"
            ElseIf optWildOn.Value = True Then
                !Wildcard = "ON"
            End If
            If optBackOFF.Value = True Then
                !BackupMX = "OFF"
            ElseIf optBackON.Value = True Then
                !BackupMX = "ON"
            End If
        .Update
    End With
    cmdClearForm_Click
End Sub

Private Sub cmdDelInfo_Click()
    rs.Delete
    MsgBox "Account Deleted from Database", vbInformation + vbOKOnly, "Deletion Complete"
    rs.MoveFirst
    cmdClearForm_Click
End Sub

Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Private Sub lstHosts_Click()
    Dim intLoopIndex
    intLoopIndex = 0
    rs.MoveFirst
    For intLoopIndex = 0 To lstHosts.ListCount - 1
        If lstHosts.Selected(intLoopIndex) Then
            Do Until rs.EOF
                If rs.Fields("Host") Like lstHosts.Text Then
                    txtHost.Text = rs.Fields("Host")
                    txtUsername.Text = DeCode(rs.Fields("Username"))
                    txtPassword.Text = DeCode(rs.Fields("Password"))
                    txtIP.Text = rs.Fields("IPAdd")
                    If rs.Fields("HostType") = "dyndns" Then
                        cmbHostType.Text = cmbHostType.List(0)
                    ElseIf rs.Fields("HostType") = "statdns" Then
                        cmbHostType.Text = cmbHostType.List(1)
                    End If
                    If rs.Fields("MXRecord") = "N/A" Then
                        txtMX.Text = ""
                    Else
                        txtMX.Text = rs.Fields("MXRecord")
                    End If
                    If rs.Fields("BackUpMX") = "OFF" Then
                        optBackOFF.Value = True
                        optBackON.Value = False
                    ElseIf rs.Fields("BackUpMX") = "ON" Then
                        optBackOFF.Value = False
                        optBackON.Value = True
                    End If
                    If rs.Fields("WildCard") = "OFF" Then
                        optWildOFF.Value = True
                        optWildOn.Value = False
                    ElseIf rs.Fields("WildCard") = "ON" Then
                        optWildOFF.Value = False
                        optWildOn.Value = True
                    End If
                    Exit Sub
                Else
                    rs.MoveNext
                End If
            Loop
        End If
    Next intLoopIndex
End Sub

Public Function DeCode(vText As String) As String
    Dim CurSpc As Integer
    Dim varLen As Integer
    Dim varChr As String
    Dim varFin As String
    CurSpc = CurSpc + 1
    varLen = Len(vText)
    Do While CurSpc <= varLen
        DoEvents
            varChr = Mid(vText, CurSpc, 3)
            Select Case varChr
                'lower case
                Case "coe"
                    varChr = "a"
                Case "wer"
                    varChr = "b"
                Case "ibq"
                    varChr = "c"
                Case "am7"
                    varChr = "d"
                Case "pm1"
                    varChr = "e"
                Case "mop"
                    varChr = "f"
                Case "9v4"
                    varChr = "g"
                Case "qu6"
                    varChr = "h"
                Case "zxc"
                    varChr = "i"
                Case "4mp"
                    varChr = "j"
                Case "f88"
                    varChr = "k"
                Case "qe2"
                    varChr = "l"
                Case "vbn"
                    varChr = "m"
                Case "qwt"
                    varChr = "n"
                Case "pl5"
                    varChr = "o"
                Case "13s"
                    varChr = "p"
                Case "c%l"
                    varChr = "q"
                Case "w$w"
                    varChr = "r"
                Case "6a@"
                    varChr = "s"
                Case "!2&"
                    varChr = "t"
                Case "(=c"
                    varChr = "u"
                Case "wvf"
                    varChr = "v"
                Case "dp0"
                    varChr = "w"
                Case "w$-"
                    varChr = "x"
                Case "vn&"
                    varChr = "y"
                Case "c*4"
                    varChr = "z"
                'numbers
                Case "aq@"
                    varChr = "1"
                Case "902"
                    varChr = "2"
                Case "2.&"
                    varChr = "3"
                Case "/w!"
                    varChr = "4"
                Case "|pq"
                    varChr = "5"
                Case "ml|"
                    varChr = "6"
                Case "t'?"
                    varChr = "7"
                Case ">^s"
                    varChr = "8"
                Case "<s^"
                    varChr = "9"
                Case ";&c"
                    varChr = "0"
                'caps
                Case "$)c"
                    varChr = "A"
                Case "-gt"
                    varChr = "B"
                Case "|p*"
                    varChr = "C"
                Case "1" & Chr(34) & "r"
                    varChr = "D"
                Case "c>:"
                    varChr = "E"
                Case "@+x"
                    varChr = "F"
                Case "v^a"
                    varChr = "G"
                Case "]eE"
                    varChr = "H"
                Case "aP0"
                    varChr = "I"
                Case "{=1"
                    varChr = "J"
                Case "cWv"
                    varChr = "K"
                Case "cDc"
                    varChr = "L"
                Case "*,!"
                    varChr = "M"
                Case "fW" & Chr(34)
                    varChr = "N"
                Case ".?T"
                    varChr = "O"
                Case "%<8"
                    varChr = "P"
                Case "@:a"
                    varChr = "Q"
                Case "&c$"
                    varChr = "R"
                Case "WnY"
                    varChr = "S"
                Case "{Sh"
                    varChr = "T"
                Case "_%M"
                    varChr = "U"
                Case "}'$"
                    varChr = "V"
                Case "QlU"
                    varChr = "W"
                Case "Im^"
                    varChr = "X"
                Case "l|P"
                    varChr = "Y"
                Case ".>#"
                    varChr = "Z"
                'Special characters
                Case "\" & Chr(34) & "]"
                    varChr = "!"
                Case "cY,"
                    varChr = "@"
                Case "x%B"
                    varChr = "#"
                Case "a*v"
                    varChr = "$"
                Case "'&T"
                    varChr = "%"
                Case ";%R"
                    varChr = "^"
                Case "eG_"
                    varChr = "&"
                Case "Z/e"
                    varChr = "*"
                Case "rG\"
                    varChr = "("
                Case "]*F"
                    varChr = ")"
                Case "@B*"
                    varChr = "_"
                Case "+Hc"
                    varChr = "-"
                Case "&|D"
                    varChr = "="
                Case "(:#"
                    varChr = "+"
                Case "SlW"
                    varChr = "["
                Case "'QB"
                    varChr = "]"
                Case "{D>"
                    varChr = "{"
                Case "+c%"
                    varChr = "}"
                Case "(s:"
                    varChr = ":"
                Case "^a("
                    varChr = ";"
                Case "16."
                    varChr = "'"
                Case "s.*"
                    varChr = Chr(34)
                Case "&?W"
                    varChr = ","
                Case "GPQ"
                    varChr = "."
                Case "SK*"
                    varChr = "<"
                Case "RL^"
                    varChr = ">"
                Case "40C"
                    varChr = "/"
                Case "?#9"
                    varChr = "?"
                Case "_?/"
                    varChr = "\"
                Case "(_@"
                    varChr = "|"
                Case "=#B"
                    varChr = " "
            End Select
        varFin = varFin & varChr
        CurSpc = CurSpc + 3
        DoEvents
    Loop
    DeCode = varFin
End Function

Public Function Encode(vText As String)
    Dim CurSpc As Integer
    Dim varLen As Integer
    Dim varChr As String
    Dim varFin As String
    varLen = Len(vText)
    Do While CurSpc <= varLen
        DoEvents
            CurSpc = CurSpc + 1
            varChr = Mid(vText, CurSpc, 1)
            Select Case varChr
                'lower case
                Case "a"
                    varChr = "coe"
                Case "b"
                    varChr = "wer"
                Case "c"
                    varChr = "ibq"
                Case "d"
                    varChr = "am7"
                Case "e"
                    varChr = "pm1"
                Case "f"
                    varChr = "mop"
                Case "g"
                    varChr = "9v4"
                Case "h"
                    varChr = "qu6"
                Case "i"
                    varChr = "zxc"
                Case "j"
                    varChr = "4mp"
                Case "k"
                    varChr = "f88"
                Case "l"
                    varChr = "qe2"
                Case "m"
                    varChr = "vbn"
                Case "n"
                    varChr = "qwt"
                Case "o"
                    varChr = "pl5"
                Case "p"
                    varChr = "13s"
                Case "q"
                    varChr = "c%l"
                Case "r"
                    varChr = "w$w"
                Case "s"
                    varChr = "6a@"
                Case "t"
                    varChr = "!2&"
                Case "u"
                    varChr = "(=c"
                Case "v"
                    varChr = "wvf"
                Case "w"
                    varChr = "dp0"
                Case "x"
                    varChr = "w$-"
                Case "y"
                    varChr = "vn&"
                Case "z"
                    varChr = "c*4"
                'numbers
                Case "1"
                    varChr = "aq@"
                Case "2"
                    varChr = "902"
                Case "3"
                    varChr = "2.&"
                Case "4"
                    varChr = "/w!"
                Case "5"
                    varChr = "|pq"
                Case "6"
                    varChr = "ml|"
                Case "7"
                    varChr = "t'?"
                Case "8"
                    varChr = ">^s"
                Case "9"
                    varChr = "<s^"
                Case "0"
                    varChr = ";&c"
                'caps
                Case "A"
                    varChr = "$)c"
                Case "B"
                    varChr = "-gt"
                Case "C"
                    varChr = "|p*"
                Case "D"
                    varChr = "1" & Chr(34) & "r"
                Case "E"
                    varChr = "c>:"
                Case "F"
                    varChr = "@+x"
                Case "G"
                    varChr = "v^a"
                Case "H"
                    varChr = "]eE"
                Case "I"
                    varChr = "aP0"
                Case "J"
                    varChr = "{=1"
                Case "K"
                    varChr = "cWv"
                Case "L"
                    varChr = "cDc"
                Case "M"
                    varChr = "*,!"
                Case "N"
                    varChr = "fW" & Chr(34)
                Case "O"
                    varChr = ".?T"
                Case "P"
                    varChr = "%<8"
                Case "Q"
                    varChr = "@:a"
                Case "R"
                    varChr = "&c$"
                Case "S"
                    varChr = "WnY"
                Case "T"
                    varChr = "{Sh"
                Case "U"
                    varChr = "_%M"
                Case "V"
                    varChr = "}'$"
                Case "W"
                    varChr = "QlU"
                Case "X"
                    varChr = "Im^"
                Case "Y"
                    varChr = "l|P"
                Case "Z"
                    varChr = ".>#"
                'Special characters
                Case "!"
                    varChr = "\" & Chr(34) & "]"
                Case "@"
                    varChr = "cY,"
                Case "#"
                    varChr = "x%B"
                Case "$"
                    varChr = "a*v"
                Case "%"
                    varChr = "'&T"
                Case "^"
                    varChr = ";%R"
                Case "&"
                    varChr = "eG_"
                Case "*"
                    varChr = "Z/e"
                Case "("
                    varChr = "rG\"
                Case ")"
                    varChr = "]*F"
                Case "_"
                    varChr = "@B*"
                Case "-"
                    varChr = "+Hc"
                Case "="
                    varChr = "&|D"
                Case "+"
                    varChr = "(:#"
                Case "["
                    varChr = "SlW"
                Case "]"
                    varChr = "'QB"
                Case "{"
                    varChr = "{D>"
                Case "}"
                    varChr = "+c%"
                Case ":"
                    varChr = "(s:"
                Case ";"
                    varChr = "^a("
                Case "'"
                    varChr = "16."
                Case Chr(34)
                    varChr = "s.*"
                Case ","
                    varChr = "&?W"
                Case "."
                    varChr = "GPQ"
                Case "<"
                    varChr = "SK*"
                Case ">"
                    varChr = "RL^"
                Case "/"
                    varChr = "40C"
                Case "?"
                    varChr = "?#9"
                Case "\"
                    varChr = "_?/"
                Case "|"
                    varChr = "(_@"
                Case " "
                    varChr = "=#B"
            End Select
        varFin = varFin & varChr
        DoEvents
    Loop
    Encode = varFin
End Function
