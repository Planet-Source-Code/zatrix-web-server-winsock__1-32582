VERSION 5.00
Object = "{B4DF8E1C-0652-41AB-9C4B-E2E04A43B747}#1.0#0"; "ZButton.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "WEB-SERVER"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5115
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin ZATRiX.Tray trIcon 
      Left            =   4050
      Top             =   990
      _ExtentX        =   529
      _ExtentY        =   529
      Icon            =   "frmMain.frx":27A2
   End
   Begin VB.Timer tmKeys 
      Interval        =   1
      Left            =   4500
      Top             =   885
   End
   Begin MSComctlLib.StatusBar stMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   3165
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   2831
            Key             =   "IP"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   1173
            MinWidth        =   882
            Key             =   "PORT"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   2055
            MinWidth        =   1764
            Key             =   "STATUS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   2831
            Key             =   "REC"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Block IPs"
      Height          =   1110
      Left            =   0
      TabIndex        =   13
      Top             =   2040
      Width           =   5055
      Begin ZButton.ButtonEx cmbUnBlck 
         Height          =   315
         Left            =   3030
         TabIndex        =   17
         Top             =   615
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   556
         Appearance      =   1
         Caption         =   "&UN-Block"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   65535
      End
      Begin ZButton.ButtonEx cmbBlock 
         Height          =   315
         Left            =   3030
         TabIndex        =   15
         Top             =   225
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   556
         Appearance      =   1
         Caption         =   "&BLOCK"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   255
      End
      Begin VB.ListBox lstHContact 
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000FF&
         Height          =   735
         ItemData        =   "frmMain.frx":6F84
         Left            =   120
         List            =   "frmMain.frx":6F86
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   225
         Width           =   2745
      End
   End
   Begin MSWinsockLib.Winsock WServer 
      Index           =   0
      Left            =   1890
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmbStop 
      Caption         =   "S&top Services"
      Enabled         =   0   'False
      Height          =   270
      Left            =   3225
      TabIndex        =   12
      ToolTipText     =   "Stop The Web Server"
      Top             =   1665
      Width           =   1815
   End
   Begin VB.CommandButton cmbStart 
      Caption         =   "&Start Services"
      Default         =   -1  'True
      Height          =   270
      Left            =   3225
      TabIndex        =   11
      ToolTipText     =   "Start The Web Server"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Connection Info"
      Height          =   870
      Left            =   0
      TabIndex        =   7
      Top             =   1155
      Width           =   3135
      Begin ZButton.ButtonEx LocalCust 
         Height          =   480
         Left            =   2325
         TabIndex        =   10
         ToolTipText     =   "Options"
         Top             =   255
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         Caption         =   ""
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentColor=   16777215
         SkinOver        =   "frmMain.frx":6F88
         SkinUp          =   "frmMain.frx":7862
         TransparentColor=   16777215
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         Caption         =   "Local Port: xxx"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   90
         TabIndex        =   9
         ToolTipText     =   "The port to which a web browser may connect to"
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label lblIp 
         AutoSize        =   -1  'True
         Caption         =   "Local IP: xxx.xxx.xxx.xxx"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   75
         TabIndex        =   8
         ToolTipText     =   "Your IP address"
         Top             =   255
         Width           =   1710
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Root Info"
      Height          =   1140
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5070
      Begin MSComDlg.CommonDialog Browser 
         Left            =   345
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Main Index File"
         FileName        =   "*.htm;*.html"
      End
      Begin VB.TextBox txtIndex 
         Height          =   300
         Left            =   900
         TabIndex        =   6
         Top             =   660
         Width           =   3360
      End
      Begin ZButton.ButtonEx BrowS 
         Height          =   480
         Left            =   4425
         TabIndex        =   4
         Top             =   555
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SkinOver        =   "frmMain.frx":A014
         SkinUp          =   "frmMain.frx":ACEE
      End
      Begin ZButton.ButtonEx BrowF 
         Height          =   480
         Left            =   4440
         TabIndex        =   3
         Top             =   120
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentColor=   12632256
         SkinDown        =   "frmMain.frx":D4A0
         SkinOver        =   "frmMain.frx":FC82
         SkinUp          =   "frmMain.frx":1095C
         TransparentColor=   12632256
      End
      Begin VB.TextBox txtRoot 
         Height          =   285
         Left            =   900
         TabIndex        =   2
         Top             =   210
         Width           =   3360
      End
      Begin VB.Label lblMainI 
         AutoSize        =   -1  'True
         Caption         =   "Main Page:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   60
         TabIndex        =   5
         Top             =   690
         Width           =   810
      End
      Begin VB.Label lblRootP 
         AutoSize        =   -1  'True
         Caption         =   "Root Path:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   75
         TabIndex        =   1
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuPort 
         Caption         =   "&Change Local Port"
      End
      Begin VB.Menu mnuCompact 
         Caption         =   "<&Compact Mode>  [Caps Lock]"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'does not allow for undeclared variables to be used
Dim TCount As Integer 'misc variables are declared |
Dim portt As Integer '                             |
Dim items_Chk As String, items_ChkU As String '    |
Dim IPs_to_be_blocked As String '                  |
Private Sub BrowF_Click() 'browse for folder
txtRoot.Text = GetFolder(Me.hwnd, "Root Folder") 'displays a promt for the root folder to be selected
End Sub

Private Sub BrowS_Click() 'the browse button for the index page
If txtRoot.Text = "" Then 'if the text field "ROOT" is empty then
MsgBox "Choose a Root Direcotory first!", vbCritical, "Error" 'display error that the user must select a directory first
Else 'otherwise
Browser.InitDir = txtRoot.Text 'chanegs the commanddialog's initial directory to the root folder
Browser.ShowOpen 'shows the open file window
txtIndex.Text = Browser.Filename 'inserts the select file path into the text field "INDEX"
txtIndex.Text = Replace(txtIndex.Text, txtRoot.Text & "\", "") 'removes the unnesessary info
End If
End Sub

Private Sub cmbUnBlck_Click() 'unblock code
Dim temp 'declares misc variables }
Dim strBlock As String '          }
Dim i As Integer '                }
Dim answ As Integer '             }
temp = Split(items_ChkU, "|") 'splits the varibale items_Chku, for "|"
For i = 0 To UBound(temp) - 1 'loops until end of array
    strBlock = lstHContact.List(temp(i)) & vbCrLf & strBlock 'creates a list of IPs to be unblocked
Next
If strBlock <> "" Then answ = MsgBox("Are you sure you want to un-block the following IP adress(es)?:" & vbCrLf & strBlock, vbCritical + vbYesNoCancel, "Block?") 'asks if the user really wants to unblock the selected IPs
If answ = vbYes Then 'if the user answers yes
    For i = 0 To UBound(temp) - 1 'loops until end of array
        lstHContact.List(temp(i)) = Replace(lstHContact.List(temp(i)), "***", "") 'removes the *** around the IPs }
        lstHContact.List(temp(i)) = Replace(lstHContact.List(temp(i)), "***", "") '                               }
        IPs_to_be_blocked = Replace(IPs_to_be_blocked, lstHContact.List(temp(i)), "") 'removes the IP from the blocked list
        lstHContact.Selected(temp(i)) = False 'unchecks the selected item
        items_Chk = Replace(items_Chk, temp(i) & "|", "") 'removes the ip from checked list
    Next
    items_ChkU = "" 'clears variable
ElseIf answ = vbNo Then 'if the user answers no then
    items_ChkU = "" 'clears variable
    For i = 0 To UBound(temp) - 1 'loops until end of array - 1
        lstHContact.Selected(temp(i)) = False 'unchecks all the selected items
    Next
Else 'else cancel
    items_ChkU = "" 'clears variable
End If
End Sub

Private Sub cmbBlock_Click() 'the block buttons code, Custom ActiveX
Dim temp ' misc variables declared }
Dim strBlock As String '           }
Dim i As Integer '                 }
Dim answ As Integer '              }
temp = Split(items_Chk, "|") 'splits the checked items varible for "|"
For i = 0 To UBound(temp) - 1 'loops until end of array
    strBlock = lstHContact.List(temp(i)) & vbCrLf & strBlock 'makes a compact string variable with all the ip addys
Next
If strBlock <> "" Then answ = MsgBox("Are you sure you want to block the following IP adress(es)?:" & vbCrLf & strBlock, vbCritical + vbYesNoCancel, "Block?") 'ask the user if he/she wants to block the selected IP addresses
If answ = vbYes Then 'if the user answers yes then
    For i = 0 To UBound(temp) - 1 'loops until end of array
        lstHContact.List(temp(i)) = "***" & lstHContact.List(temp(i)) & "***" 'adds *** around the blocked IPs to make them stand out
        IPs_to_be_blocked = lstHContact.List(temp(i)) & vbCrLf & IPs_to_be_blocked 'makes a blocked IP variable
        lstHContact.Selected(temp(i)) = False 'unchecks the current item
    Next
    items_Chk = "" 'clears the variable
ElseIf answ = vbNo Then 'if the user answers no then it simply unchecks all the items
    For i = 0 To UBound(temp) - 1 'loops until end of list
        lstHContact.Selected(temp(i)) = False 'unchecks the current list item
    Next
    items_Chk = "" 'clears the variable
Else 'if cancel then
    items_Chk = "" 'clears the variable
End If
End Sub

Private Sub cmbStart_Click()
StartServer 'starts server, calls the procedure
End Sub

Private Sub cmbStop_Click()
StopServer 'stops server, calls the procedure
End Sub

Private Sub Form_Load()
Dim inf As Integer, i As Integer 'decalres misc variables }
Dim temp, temp2 '                                         }
Dim data As String '                                      }
inf = FreeFile 'initiates variable
On Error GoTo nn 'on problem goto label
Open App.Path & "\server.ini" For Input As #inf 'opens the specified file for reading
    Do Until EOF(inf) 'does until end of file
        Line Input #inf, temp 'puts line out by line from file
        temp2 = temp2 & temp 'accumulates line together
    Loop 'loops
    temp = Split(temp2, "|") 'splits temp2 for "|"
    txtRoot = temp(1) & "\" 'text field "ROOT" is assigned the temp(1) value
    txtIndex = temp(3) ' the index file is assigned to text field "INDEX"
Close #inf 'closes file
nn: 'error label
On Error GoTo nn1 'next label error trapping
Open App.Path & "\iplog.txt" For Binary As #inf 'opens the next spcified file for binary access
    data = Space(LOF(inf)) 'fills the data file with spaces according to the amount of spaces
    Get #inf, , data 'puts the file's contents into data string
Close #inf 'closes file
Do Until InStr(data, Chr(10)) = 0 'does until chr(10) is nolonger found
    data = Replace(data, Chr(10), "") 'removes the character (10)
Loop 'repeats until conditions are met
temp = Split(data, Chr(13)) 'splits the temp variable for chr(13){linebreak}
For i = 0 To UBound(temp) 'loops until end of array
    If temp(i) <> "" Then lstHContact.AddItem temp(i) 'if temp(i) is not empty then is added to list of IPs
    If InStr(temp(i), "***") <> 0 Then IPs_to_be_blocked = temp(i) & vbCrLf & IPs_to_be_blocked 'if ip has *** around the addy then its a blocked ip addy
Next
nn1: 'error label 2
If InStr(Command, "Start") = 1 Then StartServer 'checks the string passed on start up, if "start" then automaticaly starts the server with the default settings
If InStr(Command, "Compact") = 1 Then 'if command contains "compact" then starts the server in compact mode
    SendKeys "{CAPSLOCK}" 'send the caps lock key
    Compact_Mode = True 'enables the compact mode
    Me.Hide 'hides frmmain
    StartServer 'starts server
    frmBar.Show 'shows the compact mode form (frmbar)
End If
trIcon.Visible = True 'shows an icon in the windows status bar
lblIP.Caption = "Local IP: " & WServer(0).LocalIP 'shows the local ip addy
lblPort.Caption = "Local Port: " & WServer(0).LocalPort 'shows the local port
stMain.Panels.Item(1).Text = "http://" & WServer(0).LocalIP  'displays the ip
stMain.Panels.Item(2).Text = WServer(0).LocalPort 'again displays the port
stMain.Panels.Item(3).Text = "N/A" '0 bytes
stMain.Panels.Item(4).Text = "Not Connected" 'not connected yet so this is displayed
Close 'closes all open files, incase
End Sub
Public Sub StopServer() 'stops the server procedure
cmbStop.Enabled = False 'disables the stop button
cmbStart.Enabled = True 'enables the start button
WServer(0).Close 'close the connection
End Sub

Private Sub Form_Unload(Cancel As Integer) 'when form unloads
Dim inf As Integer, i As Integer 'decalres misc variable
inf = FreeFile 'initiates the misc variable
Open App.Path & "\iplog.txt" For Output As #inf 'open the specified file for writing into it
    For i = 0 To lstHContact.ListCount - 1 'loops until end of list
        Print #inf, lstHContact.List(i) 'puts the list's ip addy into file
    Next
Close #inf 'closes file
inf = FreeFile 're-initates the variable again
Open App.Path & "\server.ini" For Output As #inf 'opens the specified file for writing
Print #inf, "[Root]|" 'prints all the required info }
Print #inf, txtRoot.Text & "|" '                    }
Print #inf, "[Index]|" '                            }
Print #inf, txtIndex.Text '                         }
Close #inf 'close the open file
End Sub

Private Sub LocalCust_Click() 'custom ActiveX
PopupMenu mnuMain 'shows the menu
End Sub

Private Sub lstHContact_ItemCheck(Item As Integer) 'when an IP addy is checked in the list box
If InStr(lstHContact.List(Item), "***") = 0 Then If lstHContact.Selected(Item) = True Then If InStr(items_Chk, Item) = 0 Then items_Chk = Item & "|" & items_Chk 'if the selected item is not blocked than add to to be blocked list
If InStr(lstHContact.List(Item), "***") <> 0 Then If lstHContact.Selected(Item) = True Then If InStr(items_ChkU, Item) = 0 Then items_ChkU = Item & "|" & items_ChkU 'if the selected item is blocked than add to to be unblocked list
End Sub

Private Sub mnuCompact_Click() 'menu Compact mode
Me.Hide 'hides frmmain
frmBar.Show 'shows frmbar
Compact_Mode = True 'changes the compact variable
End Sub

Private Sub mnuPort_Click() 'menu Port
portt = Val(InputBox("Enter a different local port.", "Port: " & WServer(0).LocalPort, "80")) 'input box asks for a port
lblPort.Caption = "Local Port: " & WServer(0).LocalPort 'shows the changed port
stMain.Panels.Item(2).Text = WServer(0).LocalPort 'shows port in status bar of form
stMain.Refresh 'refreshes form
End Sub

Private Sub tmKeys_Timer() 'timer to get the keys pressed
If GetKeyState(vbKeyCapital) = 1 Then 'checks if Caps Lock is pressed
    Me.Hide 'hide frmmain
    frmBar.Show 'show frmbar
    Compact_Mode = True 'assign true to compact variable
    tmKeys = False
End If
End Sub

Private Sub trIcon_DblClick()
PopupMenu mnuMain 'on double click on status bar icon show popup menu
End Sub

Private Sub trIcon_MouseDown(Button As Integer)
PopupMenu mnuMain 'on mouse down on status bar icon show popup menu
End Sub

Private Sub trIcon_MouseUp(Button As Integer)
PopupMenu mnuMain 'on mouse up on status bar icon show popup menu
End Sub

Private Sub WServer_Close(Index As Integer)
WServer(Index).Close 'on connection close, close the connection, makes sense eh!
End Sub

Private Sub WServer_Connect(Index As Integer)
If Compact_Mode Then Compact_Display 2 * Rnd 'changes the status icons in frmbar
End Sub

Private Sub WServer_ConnectionRequest(Index As Integer, ByVal requestID As Long) 'if a client browser is requesting a connection
    If Index = 0 Then 'if winsock 0
        If Compact_Mode Then Compact_Display 2 * Rnd 'changes the status icons in frmbar
        frmBar.lblIP.Caption = WServer(0).RemoteHostIP & " " & WServer(0).RemotePort 'in frmbar shows the current remote host and the port of connection
        TCount = TCount + 1 'makes more connection available
        Load WServer(TCount) 'loads another instance of winsock
        WServer(TCount).LocalPort = 0 'clears local port for loaded winsock
        WServer(TCount).Accept requestID 'accepts connection
        stMain.Panels.Item(4).Text = WServer(0).RemoteHostIP 'displays current connection (IP)
    End If
End Sub
Public Function IpBlockable(ByVal ip As String) As Integer 'function to add the current connection to list
If IpExistance(WServer(0).RemoteHostIP) = 0 Then 'if current ip does not exist in list then
                lstHContact.AddItem WServer(0).RemoteHostIP 'adds the ip to list
            Else 'otherwise
                If InStr(IPs_to_be_blocked, WServer(0).RemoteHostIP) <> 0 Then IpBlockable = 1 Else IpBlockable = 0 'return either 0 or 1 (false or true)
End If
End Function
Public Function IpExistance(ByVal ip As String) As Integer 'functions to see if ip address exists in list
Dim i As Integer 'misc variable
For i = 0 To lstHContact.ListCount - 1 'loops until the end of list
    If InStr(lstHContact.List(i), ip) <> 0 Then 'if ip addy is found then
        IpExistance = 1 ' return true
        i = lstHContact.ListCount - 1 'and quit loop
    Else 'otherwise
        IpExistance = 0 'return false
    End If
Next
End Function
Public Sub Compact_Display(action As Integer) 'procedure which changes the status icons in frmbar
If action = 0 Then
    frmBar.imgStatus.Picture = frmBar.IconLst.ListImages(1).Picture
Else
    frmBar.imgStatus.Picture = frmBar.IconLst.ListImages(2).Picture
End If
End Sub
Private Sub WServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String 'declares misc variable            }
Dim StrGet As Integer, spc2 As Integer, i As Integer '   }
Dim Page As String '                                     }
Dim DataRes, DataRes2 '                                  }  all together
Dim data As String '                                     }
Dim Found As Boolean '                                   }
Dim username As String, pass As String '                 }
If Compact_Mode Then Compact_Display 2 * Rnd 'changes the status icons in frmbar
    stMain.Panels.Item(3).Text = bytesTotal 'displays the size of incomming message
    WServer(Index).GetData strData, vbString 'gets the sent data by client
    If Mid(strData, 1, 3) = "GET" Then 'if browser is trying to get a page do the below
        StrGet = InStr(strData, "GET ") 'tries to find "GET"
        spc2 = InStr(StrGet + 5, strData, " ") 'attempsts to find a space character
        Page = Trim(Mid(strData, StrGet + 5, spc2 - (StrGet + 4))) 'trims to find the requested file
        If Right(Page, 1) = "/" Then Page = Left(Page, Len(Page) - 1) 'does misc calculations
        SendData Page, Index 'calls a procedure which will send the specified file back to the visitor
    ElseIf Mid(strData, 1, 5) = "Uname" Or Mid(strData, 1, 4) = "POST" Then 'if the browser is posting then starts the logging in prossess
        On Error GoTo errr 'if a problem occurs goto label
        StrGet = InStr(strData, "=") 'finds the first instace of "="
        spc2 = InStr(strData, "&") 'finds the first instance of "&"
        username = Mid(strData, StrGet + 1, spc2 - StrGet - 1) 'the username is extracted from a string
        StrGet = InStr(strData, "s=") 'strGet is assigned the frist instance of "s="
        spc2 = InStr(strData, "&S") 'spc2 is assigned the first instance of the comdination "&S"
        pass = Mid(strData, StrGet + 2, spc2 - StrGet - 2) 'takes the specified characters from string to be a password
errr: 'an error label
        Dim inf As Integer 'Declares a required integer
        inf = FreeFile 'makes the variable a FREEFILE
        Open txtRoot.Text & "dataFi.bin" For Binary As #inf 'opens the specified file for editing
            data = Space(LOF(inf)) 'fills the string with space according to length of file
            Get #inf, , data 'puts the contents of file into string
        Close #inf 'closes opened file
        DataRes = Split(data, Chr(13))
        For i = 1 To UBound(DataRes) 'loops until end of array
            DataRes2 = Split(DataRes(i), ",") 'splits Datares for commas
            If username = DataRes2(1) And pass = DataRes2(2) Then 'check to see if the username and password can be found in database file
                i = UBound(DataRes) 'exits the for loop
                WServer(Index).SendData FoundUser(DataRes2(0)) 'calls a function which return an appropriate file
                Found = True 'found
            Else
                Found = False 'not found
            End If
        Next
            If Not Found Then SendData "hack.html", Index 'checks if name and password were found in database file
    End If
Page = "" 'on exit of sub clears variables of unwanted data }
strData = "" '                                              }
End Sub
Sub SendData(Page As String, Index As Integer)
Dim data As String 'declares data string
Dim hFile As Long 'declares  hFile integer
On Error Resume Next
    hFile = FreeFile 'makes the file hFile available for work with
    If IpBlockable(WServer(0).RemoteHostIP) = 1 Then Page = "blocked.html" 'checks if the current computer is blocked, if is than does not let computer see the default page
    If Page = "" Then
        Page = txtIndex.Text 'if page equals nothing than makes it equall the start up page
    ElseIf Page = "!dbase.exe!" Then
        Page = "dataFi.bin" 'return the database file
    End If
    If Not FileExists(txtRoot.Text & Page) Then Page = "404.html" 'checks if file exists in root directory
    Open txtRoot.Text & Page For Binary As #hFile ' open the specified file for binary work
         data = Space(LOF(hFile)) 'fills the string with spaces according to the length of file
         frmBar.pb.Max = LOF(hFile) 'progress bar gets the size of file assigned to it
         frmBar.lblRequest.Caption = Page 'in frmbar shows the requested item
         Get #hFile, , data 'puts the whole file into variable
    Close #hFile 'closes the open file
    WServer(Index).SendData data 'send back the contents of DATA
    If Compact_Mode Then Compact_Display 2 * Rnd 'changes the icons in frmbar
End Sub
Public Sub StartServer()
On Error GoTo SeverErr 'if a problem occurs goto label
lblPort.Caption = "Local Port: " & WServer(0).LocalPort 'shows the current connection port
stMain.Panels.Item(2).Text = WServer(0).LocalPort 'shows the current connection port
cmbStart.Enabled = False 'disables the start button
cmbStop.Enabled = True 'enables the stop button
SeverErr: 'an error goto label
    If Err.Number = 10048 Then 'in case there is a problem it displays the message
        MsgBox "Another service is in use please close down any other services you may have open and try agian.", vbInformation
        Exit Sub
    Else
        WServer(0).Close 'closes winsock
        If portt = 0 Then portt = 80 'if the port is 0 then changes it to 80 (default)
        WServer(0).LocalPort = portt 'assigns the port
        WServer(0).Listen 'listen for a request
    End If
    
End Sub

Private Sub WServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If Compact_Mode Then Compact_Display 2 * Rnd 'changes the status icons in frmbar
End Sub

Private Sub WServer_SendComplete(Index As Integer)
If Compact_Mode Then Compact_Display 2 * Rnd 'changes icons in frmbar
WServer(Index).Close 'closes the current connection
End Sub

Private Sub WServer_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
On Error Resume Next 'if a problem happens continue anyway
If Compact_Mode Then Compact_Display 2 * Rnd 'changes the bars icons around
frmBar.pb.Value = bytesRemaining 'creates a progress status bar
End Sub
