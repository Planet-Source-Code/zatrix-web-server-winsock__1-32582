VERSION 5.00
Object = "{B4DF8E1C-0652-41AB-9C4B-E2E04A43B747}#1.0#0"; "ZButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBar 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmBar.frx":0000
   ScaleHeight     =   600
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ZButton.ButtonEx lblClose 
      Height          =   240
      Left            =   4350
      TabIndex        =   5
      Top             =   60
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   423
      Appearance      =   0
      BackColor       =   7893617
      Caption         =   "X"
      ForeColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   16777215
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   180
      Left            =   900
      TabIndex        =   2
      Top             =   105
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList IconLst 
      Left            =   1665
      Top             =   630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBar.frx":0839
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBar.frx":0B53
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmSlideI 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1185
      Top             =   690
   End
   Begin VB.Timer tmSlideO 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   720
      Top             =   675
   End
   Begin VB.PictureBox Maskk 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   60
      Picture         =   "frmBar.frx":142D
      ScaleHeight     =   570
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   630
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblRequest 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "xxx.html"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2490
      TabIndex        =   4
      Top             =   330
      Width           =   2100
   End
   Begin VB.Label lblIP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "xxx.xxx.xxx.xxx"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   780
      TabIndex        =   3
      Top             =   300
      Width           =   1725
   End
   Begin VB.Image imgStatus 
      Height          =   480
      Left            =   285
      Picture         =   "frmBar.frx":A6B1
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblSlide 
      BackStyle       =   0  'Transparent
      Height          =   585
      Left            =   0
      MousePointer    =   9  'Size W E
      TabIndex        =   1
      Top             =   0
      Width           =   345
   End
End
Attribute VB_Name = "frmBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'does not allow for undeclared variables to be used
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long 'misc Api declarations |
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long '                                               |
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long '                                                                                                 |
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long '                                                                       |
Private Declare Function ReleaseCapture Lib "user32" () As Long '                                                                                                                   |
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long '                      |
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long '                                                     |

Private Const RGN_OR = 2 'misc constants
Private lngRegion As Long 'private variable
Dim slide_Ctr As Integer 'misc variables }
Dim S_Out As Boolean, S_In As Boolean '  }
Function RegionFromBitmap(picSource As PictureBox, Optional lngTransColor As Long) As Long 'complex function to remove pixels from a form
  Dim lngRetr As Long, lngHeight As Long, lngWidth As Long
  Dim lngRgnFinal As Long, lngRgnTmp As Long
  Dim lngStart As Long, lngRow As Long
  Dim lngCol As Long
  If lngTransColor& < 1 Then
    lngTransColor& = GetPixel(picSource.hDC, 0, 0)
  End If
  lngHeight& = picSource.Height / Screen.TwipsPerPixelY
  lngWidth& = picSource.Width / Screen.TwipsPerPixelX
  lngRgnFinal& = CreateRectRgn(0, 0, 0, 0)
  For lngRow& = 0 To lngHeight& - 1
    lngCol& = 0
    Do While lngCol& < lngWidth&
      Do While lngCol& < lngWidth& And GetPixel(picSource.hDC, lngCol&, lngRow&) = lngTransColor&
        lngCol& = lngCol& + 1
      Loop
      If lngCol& < lngWidth& Then
        lngStart& = lngCol&
        Do While lngCol& < lngWidth& And GetPixel(picSource.hDC, lngCol&, lngRow&) <> lngTransColor&
          lngCol& = lngCol& + 1
        Loop
        If lngCol& > lngWidth& Then lngCol& = lngWidth&
        lngRgnTmp& = CreateRectRgn(lngStart&, lngRow&, lngCol&, lngRow& + 1)
        lngRetr& = CombineRgn(lngRgnFinal&, lngRgnFinal&, lngRgnTmp&, RGN_OR)
        DeleteObject (lngRgnTmp&)
      End If
    Loop
  Next
  RegionFromBitmap& = lngRgnFinal&
End Function
Sub ChangeMask()
On Error Resume Next ' In case of error
' To update if the skin is changed
  Dim lngRetr As Long
  lngRegion& = RegionFromBitmap(Maskk, vbWhite)
  lngRetr& = SetWindowRgn(Me.hwnd, lngRegion&, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) 'keydown procedure
Select Case KeyCode 'creates a case for KeyCode
Case vbKeyCapital 'if Caps Lock
    Unload Me 'close frmbar
    frmMain.Show 'load frmmain
    frmMain.tmKeys = True 'enable the timer in frmmain
Case vbKeyAdd '+
    tmSlideO = True 'timer slide out is enabled
    tmSlideI = False 'timer slide in is disabled
Case vbKeySubtract '-
    tmSlideO = False 'timer slide out is disabled
    tmSlideI = True 'timer slide in is enabled
Case vbKeyDelete 'Delete
    lblClose_Click 'call the close button
End Select
End Sub

Private Sub Form_Load() 'on form load
Me.Left = Screen.Width - (imgStatus.Width + lblSlide.Width) 'changes the left location from top left corner (0,0)
Me.Top = Screen.Height - (Me.Height * 2) 'changes the location from top
Maskk.AutoSize = True 'enlarges the mask for transperancy
lblSlide.ZOrder 1 'the slide label is brought to top
Call ChangeMask 'calls the change of mask procedure
AlwaysOnTop "Enabled", Me 'enables the always on top
End Sub

Private Sub Form_Unload(Cancel As Integer) 'when frmbar is unloaded
frmMain.trIcon.Visible = False 'removes the tray icon
End Sub

Private Sub lblClose_Click() 'when the close button is clicked
Dim rep 'misc variable
rep = MsgBoxEx("Do you wish to exit completely or return to full mode?", vbExclamation + vbOKCancel, "", , , False, Me.hwnd, "Full Mode", "Exit", "") 'custom msgbox with custom buttons, asking if the user wants to quit or return to full mode
If rep = vbOK Then 'if the user says Full Mode then
    Unload Me 'unload me
    Compact_Mode = False 'changes the compact variable to false
    frmMain.Show 'shows the main form (frmmain)
Else 'else if quit
    Unload Me 'simply quit
    End 'end the runtime program completely
End If
End Sub

Private Sub lblSlide_Click() 'the slide label
If S_Out Then 'if the form is already out then
    tmSlideI = True 'scroll in
Else
    tmSlideO = True 'scroll out
End If
End Sub

Private Sub tmSlideI_Timer() 'the slide in timer
slide_Ctr = slide_Ctr + 1.5 'determines the slide rate
If Me.Left >= Screen.Width - (imgStatus.Width + lblSlide.Width) Then 'stops the sliding at a certain frmbar.left
    tmSlideI = False 'disables itself
    S_In = True 'assigns true to the s_in variable
    S_Out = False 'assigns false to the s_out variable
    slide_Ctr = 0 'clears the rate
Else
    Me.Left = Me.Left + slide_Ctr 'continues to slide the form in
End If
End Sub

Private Sub tmSlideO_Timer() 'timer to slide the form out
slide_Ctr = slide_Ctr + 1.5 'determines the slide rate
If Me.Left + 51 <= (Screen.Width - Me.Width) Then 'if the form is out enough than
    tmSlideO = False 'stop timer (self)
    S_Out = True 'enable the variable
    S_In = False 'disable the variable
    slide_Ctr = 0 'clears the slide rate
Else
    Me.Left = Me.Left - slide_Ctr 'continue to slide
End If
End Sub
