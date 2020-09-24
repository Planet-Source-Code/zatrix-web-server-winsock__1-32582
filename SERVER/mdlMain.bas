Attribute VB_Name = "Main"
Global Compact_Mode As Boolean
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
Function GetFolder(ByVal hWndOwner As Long, ByVal sTitle As String) As String
Dim bInf As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim RetPath As String
Dim Offset As Integer
    bInf.hOwner = hWndOwner
    bInf.lpszTitle = sTitle
    bInf.ulFlags = BIF_RETURNONLYFSDIRS
    PathID = SHBrowseForFolder(bInf)
    RetPath = Space$(512)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
    If RetVal Then
      Offset = InStr(RetPath, Chr$(0))
      GetFolder = Left$(RetPath, Offset - 1)
    End If
End Function
Public Function LoadFile(filename1 As String) As String
On Error GoTo hell
Open filename1 For Binary As #1
LoadFile = Input(FileLen(filename1), #1)
Close #1
hell:
If Err.Number = 76 Then LoadFile = "Cant find file! Oh no!"
End Function
Public Function FileExists(ByVal Filename As String) As Boolean
    If Dir(Filename) = "" Then FileExists = False Else FileExists = True
End Function
Public Function FoundUser(ByVal name As String) As String
Dim ret As String
ret = "<html>"
ret = ret & "<head>"
ret = ret & "<title>Admin: " & name & "</title>"
ret = ret & "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
ret = ret & "</head>"
ret = ret & "<body bgcolor='#FFFFFF' text='#000000'>"
ret = ret & "<div align='center'>"
ret = ret & "<table width='75%' border='2' bordercolor='#000000' bgcolor='#00CCFF'>"
ret = ret & "<tr>"
ret = ret & "<td>"
ret = ret & "<div align='center'><b><font face='Geneva, Arial, Helvetica, san-serif' size='6'>Welcome " & name & " Administrator!</font></b></div>"
ret = ret & "</td>"
ret = ret & "</tr>"
ret = ret & "</table>"
ret = ret & "</div>"
ret = ret & "</body>"
ret = ret & "</html>"
FoundUser = ret
End Function
Public Sub AlwaysOnTop(EnabledOrDisabled, FormID As Object)
     If EnabledOrDisabled = "Enabled" Then SetWindowPos FormID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
     If EnabledOrDisabled = "Disabled" Then SetWindowPos FormID.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
