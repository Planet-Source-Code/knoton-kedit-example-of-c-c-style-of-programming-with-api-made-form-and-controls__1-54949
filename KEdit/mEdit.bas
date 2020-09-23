Attribute VB_Name = "mEdit"
Option Explicit
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

Private Type NMHDR
    hwndFrom As Long
    idFrom As Long
    code As Long
End Type

Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type


Private Type ENLINK
    hdr As NMHDR
    msg As Long
    wParam As Long
    lParam As Long
    chrg As CHARRANGE
End Type

'Menu item structure
Public Type MENUITEMINFO
    cbSize                                  As Long
    fMask                                   As Long
    fType                                   As Long
    fState                                  As Long
    wID                                     As Long
    hSubMenu                                As Long
    hbmpChecked                             As Long
    hbmpUnchecked                           As Long
    dwItemData                              As Long
    dwTypeData                              As String
    cch                                     As Long
End Type

Public Enum RE_TextModeFlags
    TM_PLAINTEXT = 1
    TM_RICHTEXT = 2             ' default behavior
    TM_SINGLELEVELUNDO = 4
    TM_MULTILEVELUNDO = 8       ' default behavior
    TM_SINGLECODEPAGE = 16
    TM_MULTICODEPAGE = 32       ' default behavior
End Enum


'// General messages
Const WM_NULL = &H0
Const WM_CREATE = &H1
Const WM_DESTROY = &H2
Const WM_SETTEXT = &HC
Const WM_GETTEXT = &HD
Const WM_GETTEXTLENGTH = &HE
Const WM_QUIT = &H12
Const WM_SETCURSOR = &H20
Public Const WM_SETFONT = &H30
Const WM_GETFONT = &H31
Const WM_GETOBJECT = &H3D
Const WM_COPYDATA = &H4A
Const WM_NOTIFY = &H4E
Const WM_HELP = &H53
Const WM_NOTIFYFORMAT = &H55
Const WM_SIZE = &H5
Const WM_SIZING = &H214
Const WM_COMMAND = &H111
Const WM_EXITSIZEMOVE = &H232
Const WM_LBUTTONDBLCLK = &H203
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_MOUSEMOVE = &H200
Const WM_RBUTTONDBLCLK = &H206
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_HOTKEY = &H312
Const WM_PARENTNOTIFY = &H210

 '// General window styles
Const WS_OVERLAPPED = &H0&
Const WS_POPUP = &H80000000
Const WS_CHILD = &H40000000
Const WS_MINIMIZE = &H20000000
Const WS_VISIBLE = &H10000000
Const WS_DISABLED = &H8000000
Const WS_CLIPSIBLINGS = &H4000000
Const WS_CLIPCHILDREN = &H2000000
Const WS_MAXIMIZE = &H1000000
Const WS_CAPTION = &HC00000
Const WS_BORDER = &H800000
Const WS_DLGFRAME = &H400000
Const WS_VSCROLL = &H200000
Const WS_HSCROLL = &H100000
Const WS_SYSMENU = &H80000
Const WS_THICKFRAME = &H40000
Const WS_GROUP = &H20000
Const WS_TABSTOP = &H10000
Const WS_MINIMIZEBOX = &H20000
Const WS_MAXIMIZEBOX = &H10000
Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Const WS_EX_ACCEPTFILES = &H10&

'// Edit Control Styles
Const ES_LEFT = &H0&
Const ES_CENTER = &H1&
Const ES_RIGHT = &H2&
Const ES_MULTILINE = &H4&
Const ES_UPPERCASE = &H8&
Const ES_LOWERCASE = &H10&
Const ES_PASSWORD = &H20&
Const ES_AUTOVSCROLL = &H40&
Const ES_AUTOHSCROLL = &H80&
Const ES_NOHIDESEL = &H100&
Const ES_OEMCONVERT = &H400&
Const ES_READONLY = &H800&
Const ES_WANTRETURN = &H1000&
Const ES_NUMBER = &H2000&

Const WM_USER = &H400
Const GWL_WNDPROC = -4

Const OCM__BASE = (WM_USER + &H1C00)
Const OCM_NOTIFY = (OCM__BASE + WM_NOTIFY)

'// Event Masks
Const ENM_NONE = &H0
Const ENM_CHANGE = &H1
Const ENM_UPDATE = &H2
Const ENM_SCROLL = &H4
Const ENM_KEYEVENTS = &H10000
Const ENM_MOUSEEVENTS = &H20000
Const ENM_REQUESTRESIZE = &H40000
Const ENM_SELCHANGE = &H80000
Const ENM_DROPFILES = &H100000
Const ENM_PROTECTED = &H200000
Const ENM_CORRECTTEXT = &H400000               ' /* PenWin specific */
Const ENM_SCROLLEVENTS = &H8
Const ENM_DRAGDROPDONE = &H10

Const EM_SETTARGETDEVICE = (WM_USER + 72)
Const EM_SETTEXTMODE = (WM_USER + 89)
Const EM_EXLIMITTEXT = (WM_USER + 53)
Const EM_SETEVENTMASK = (WM_USER + 69)
Const EM_GETEVENTMASK = (WM_USER + 59)
Const EN_MSGFILTER = &H700
Const EM_SETUNDOLIMIT = (WM_USER + 82)
Const EM_REDO = (WM_USER + 84)
Const EN_LINK = &H70B&
Const EM_SETCHARFORMAT = (WM_USER + 68)
Const EM_SETFONTSIZE = (WM_USER + 223)

'Menu constants
Public Const MIIM_BITMAP                  As Long = &H80
Public Const MIIM_CHECKMARKS              As Long = &H8
Public Const MIIM_DATA                    As Long = &H20
Public Const MIIM_FTYPE                   As Long = &H100
Public Const MIIM_ID                      As Long = &H2
Public Const MIIM_STATE                   As Long = &H1
Public Const MIIM_STRING                  As Long = &H40
Public Const MIIM_SUBMENU                 As Long = &H4
Public Const MIIM_TYPE                    As Long = &H10
Public Const MFT_MENUBARBREAK             As Long = &H20&
Public Const MFT_MENUBREAK                As Long = &H40&
Public Const MFT_OWNERDRAW                As Long = &H100&
Public Const MFT_RADIOCHECK               As Long = &H200&
Public Const MFT_RIGHTJUSTIFY             As Long = &H4000&
Public Const MFT_RIGHTORDER               As Long = &H2000&
Public Const MFT_SEPARATOR                As Long = &H800&
Public Const MFS_CHECKED                  As Long = &H8&
Public Const MFS_DEFAULT                  As Long = &H1000&
Public Const MFS_DISABLED                 As Long = &H3&
Public Const MFS_ENABLED                  As Long = &H0&
Public Const MFS_HILITE                   As Long = &H80&
Public Const MFS_UNCHECKED                As Long = &H0&
Public Const MFS_UNHILITE                 As Long = &H0&
Const MF_APPEND = &H100&
Const MF_STRING = &H0&

Const MOD_CTRL = &H2
Const IMAGE_ICON = 1
Const LR_DEFAULTCOLOR = &H0&
Const WM_MOVING = &H2
Const WM_WINDOWPOSCHANGED = &H47

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long

'Send messages
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub PostQuitMessage Lib "user32.dll" (ByVal nExitCode As Long)
Private Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Message loop functions
Private Declare Function TranslateMessage Lib "user32.dll" (lpMsg As msg) As Long
Private Declare Function GetMessage Lib "user32.dll" Alias "GetMessageA" (lpMsg As msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function DispatchMessage Lib "user32.dll" Alias "DispatchMessageA" (lpMsg As msg) As Long

'Create windows
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'Edit,refresh windows
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function UpdateWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function SetFocus Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

'Menues
Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function CreateMenu Lib "user32.dll" () As Long
Private Declare Function SetMenu Lib "user32.dll" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Private Declare Function CheckMenuItem Lib "user32.dll" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpmii As MENUITEMINFO) As Long

'Hotkeys
Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long) As Long


Public m_hWnd As Long
Public Edit_hWind As Long
Private lngLIB As Long
Private glWinRet As Long
Private menu_Hwnd As Long
Private FileMenu_Hwnd As Long
Private EditMenu_Hwnd As Long
Private CD As CommonDialog
Private CurFile As String
Private CurText As String
Private blnCheckad As Boolean

'Main sub, initalize and start messageloop
Sub Main()
Dim aMsg As msg, temp As Long, i As Integer
Set CD = New CommonDialog
'create form, menues, richedit
CreateForm
InitMenu
LoadEdit

'create hotkeys
InitHotKey

'Init fonttype
mFontType.sName = "Arial"
mFontType.Size = 12
mFontType.mfont = FW_NORMAL

'Our messageloop
Do While GetMessage(aMsg, 0, 0, 0)
    TranslateMessage aMsg
    DispatchMessage aMsg
Loop

'End of execution of our program do som cleaning up
For i = 0 To 5
    UnregisterHotKey m_hWnd, i
Next

UnregisterClass "KEdit", App.hInstance

End Sub

'common dialog open
Private Sub OpenFile()
With CD
    .DialogTitle = "Select file to open"
    .Filter = "Text Files|*.txt;*.doc;*.rtf;*.nfo;|All Files (*.*)|*.*"
    .ShowOpen
    CurFile = .Filename
    If CurFile <> "" Then Text ReadFile(CurFile)
End With
SetFocus Edit_hWind
End Sub

'Open a textfile and return the content
Public Function ReadFile(ByVal sFileName As String) As String
Dim fhFile As Integer
fhFile = FreeFile
Open sFileName For Binary As #fhFile
ReadFile = Input$(LOF(fhFile), fhFile)
Close #fhFile
End Function

'Set the text of the richedit
Private Sub Text(ByVal strText As String)
Call SendMessage(Edit_hWind, WM_SETTEXT, 0, ByVal strText)
End Sub

'Common dialog save and write to file
Private Sub WriteFile()
Dim intFilNr As Integer
Dim sFile As Variant
Dim WriteWhat As String

CD.Filter = "Text Files|*.txt;*.doc;*.rtf;*.nfo;|All Files (*.*)|*.*"
CD.DialogTitle = "Save Text File"
CD.AllowMultiSelect = False
CD.ShowSave
sFile = CD.Filename

If sFile <> "" Then
    WriteWhat = GetText
    intFilNr = FreeFile
    Open sFile For Append As intFilNr
    Print #intFilNr, WriteWhat
    Close #intFilNr
    CurFile = sFile
End If

End Sub

'Save current text in richedit to the file we opened
Private Sub SaveToFile()
Dim intFilNr As Integer
Dim WriteWhat As String

WriteWhat = GetText
intFilNr = FreeFile
Open CurFile For Output As intFilNr
Print #intFilNr, WriteWhat
Close #intFilNr

End Sub
'Initialize our hotkeys
Private Sub InitHotKey()
Dim ret As Long
ret = RegisterHotKey(m_hWnd, 0, MOD_CTRL, vbKeyO)
ret = RegisterHotKey(m_hWnd, 1, MOD_CTRL, vbKeyS)
ret = RegisterHotKey(m_hWnd, 2, MOD_CTRL, vbKeyX)
ret = RegisterHotKey(m_hWnd, 3, MOD_CTRL, vbKeyF)
ret = RegisterHotKey(m_hWnd, 4, MOD_CTRL, vbKeyA)
ret = RegisterHotKey(m_hWnd, 5, MOD_CTRL, vbKeyN)

End Sub

'create our main form
Private Sub CreateForm()
Dim wc As WNDCLASS
With wc
    .lpfnwndproc = GetAdd(AddressOf WndProc)
    .hbrBackground = 5
    .lpszClassName = "KEdit"
    .hIcon = LoadImage(App.hInstance, "101", IMAGE_ICON, 0, 0, 0)
End With

RegisterClass wc
m_hWnd = CreateWindowEx(0&, "KEdit", "KEdit", WS_OVERLAPPEDWINDOW Or WS_THICKFRAME, 150, 150, 708, 528, 0, 0, App.hInstance, ByVal 0&)
ShowWindow m_hWnd, 1
UpdateWindow m_hWnd
SetFocus m_hWnd
End Sub

'get the current text in the richedit
Private Function GetText() As String

    Dim sBuffer As String
    Dim lLen As String
    
    lLen = SendMessage(Edit_hWind, WM_GETTEXTLENGTH, 0, ByVal 0&)
    
    If lLen > 0 Then
        sBuffer = String$(lLen, vbNullChar)
        Call SendMessage(Edit_hWind, WM_GETTEXT, lLen, ByVal sBuffer)
    
    End If
    
    GetText = sBuffer

End Function

'Create or richedit
Private Sub LoadEdit()
Dim style As Long, dwMask As Long, txtMode As Long
lngLIB = LoadLibrary("RICHED20.DLL")
dwMask = ENM_MOUSEEVENTS


style = WS_CHILD Or WS_VISIBLE Or ES_MULTILINE Or WS_VSCROLL Or WS_HSCROLL Or _
        ES_AUTOHSCROLL Or ES_AUTOVSCROLL

Edit_hWind = CreateWindowEx(0, "RichEdit20a", "KEdit", _
                        style, _
                        1, 1, 700, 500, _
                        m_hWnd, 0, App.hInstance, ByVal 0&)
                        
Call SendMessage(Edit_hWind, EM_SETEVENTMASK, 0, ByVal dwMask)

ShowWindow Edit_hWind, 1
UpdateWindow Edit_hWind
SetFocus Edit_hWind
Text ""
SetTextMode TM_PLAINTEXT
SetMaxText
SetUndoLimit
ResizeEdit
End Sub

'which textmode our richedit should use, we use plain text
Private Sub SetTextMode(ByVal Tmode As RE_TextModeFlags)
Call SendMessage(Edit_hWind, EM_SETTEXTMODE, ByVal Tmode, 0&)
End Sub

'set the max lenght of our richedit, we use indefinate lenght
Private Sub SetMaxText(Optional ByVal TLen As Long = 0)
Call SendMessage(Edit_hWind, EM_EXLIMITTEXT, ByVal TLen, 0&)
End Sub

'set no of undo possible in our richedit
Private Sub SetUndoLimit(Optional ByVal TUndoLimit As Long = 3)
Call SendMessage(Edit_hWind, EM_SETUNDOLIMIT, ByVal TUndoLimit, 0)
End Sub

'Create our menues
Private Sub InitMenu()
Dim mnuItem As MENUITEMINFO

'    Create the File Sub Menu
FileMenu_Hwnd = CreatePopupMenu()
AppendMenu FileMenu_Hwnd, MF_STRING, ByVal 106, "New" & vbTab & "Ctrl+N"
AppendMenu FileMenu_Hwnd, MF_STRING, ByVal 101, "Open" & vbTab & "Ctrl+O"
AppendMenu FileMenu_Hwnd, MF_STRING, ByVal 102, "Save As" & vbTab & "Ctrl+A"
AppendMenu FileMenu_Hwnd, MF_STRING, ByVal 105, "Save" & vbTab & "Ctrl+S"

AppendMenu FileMenu_Hwnd, MFT_SEPARATOR, ByVal 0, ""
AppendMenu FileMenu_Hwnd, MF_STRING, ByVal 103, "Exit" & vbTab & "Ctrl+X"

'    Create the Edit Sub Menu
EditMenu_Hwnd = CreatePopupMenu()
AppendMenu EditMenu_Hwnd, MF_STRING, ByVal 104, "Font" & vbTab & "Ctrl+F"

'    Create the main menu
menu_Hwnd = CreateMenu()
mnuItem = CreateMenuItem("File", 1, FileMenu_Hwnd)
InsertMenuItem menu_Hwnd, 1, False, mnuItem
mnuItem = CreateMenuItem("Edit", 2, EditMenu_Hwnd)
InsertMenuItem menu_Hwnd, 2, False, mnuItem
SetMenu m_hWnd, menu_Hwnd

End Sub

'Set menue options
Public Function CreateMenuItem(ByVal strCaption As String, ByVal ID As Integer, ByVal hSubMenu As Long, _
                               Optional ByVal blnSeparator As Boolean = False, Optional ByVal blnDisabled As Boolean = False, _
                               Optional ByVal blnChecked As Boolean = False) As MENUITEMINFO

Dim mnuInfo As MENUITEMINFO

With mnuInfo
    .cbSize = Len(mnuInfo)
    .fMask = MIIM_ID Or MIIM_STRING
    If hSubMenu Then
        .fMask = .fMask Or MIIM_SUBMENU
        .hSubMenu = hSubMenu
    End If
    If blnSeparator Then
        .fMask = .fMask Or MIIM_FTYPE
        .fType = MFT_SEPARATOR
    End If
    If blnDisabled Or blnChecked Then
        .fMask = .fMask Or MIIM_STATE
        If blnDisabled Then
            .fState = MFS_DISABLED
        Else
            .fState = MFS_CHECKED
        End If
    End If
    .wID = ID
    .dwTypeData = strCaption
    .cch = Len(.dwTypeData)
End With
CreateMenuItem = mnuInfo

End Function

'Workaround for adressof
Private Function GetAdd(Address As Long) As Long
GetAdd = Address
End Function

'Resize our richedit window
Private Function ResizeEdit()
Dim R As RECT
GetWindowRect m_hWnd, R
MoveWindow Edit_hWind, 0, 0, R.Right - R.Left - 8, R.Bottom - R.Top - 46, Abs(True)
SetFocus Edit_hWind
End Function

'Return the high word of a long value.
Public Function GetHiWord(ByVal Value As Long) As Integer
CopyMemory GetHiWord, ByVal VarPtr(Value) + 2, 2
End Function

'Return the low word of a long value
Public Function GetLoWord(ByVal Value As Long) As Integer
CopyMemory GetLoWord, Value, 2
End Function

'Our window message handler
Public Function WndProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim nmh As NMHDR, eLink As ENLINK

WndProc = 0

Select Case wMsg
    Case WM_COMMAND
        If GetLoWord(wParam) = 101 Then OpenFile
        If GetLoWord(wParam) = 102 Then WriteFile
        If GetLoWord(wParam) = 103 Then PostQuitMessage 0
        If GetLoWord(wParam) = 104 Then ShowFont FW_NORMAL, False, False
        If GetLoWord(wParam) = 105 Then
            If CurFile <> "" Then
                SaveToFile
            Else
                WriteFile
            End If
        End If
        If GetLoWord(wParam) = 106 Then
            Text ""
            CurFile = ""
        End If
    Case WM_SIZE, WM_SIZING
        DoEvents
        ResizeEdit
    Case WM_DESTROY
        PostQuitMessage 0
    Case WM_HOTKEY
        Select Case wParam
            Case 0
                OpenFile
            Case 1
                If CurFile <> "" Then
                    SaveToFile
                Else
                    WriteFile
                End If
            Case 2
                PostQuitMessage 0
            Case 3
                ShowFont FW_NORMAL, False, False
            Case 4
                WriteFile
            Case 5
                Text ""
                CurFile = ""
        End Select
    Case Else
        WndProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
End Select
End Function
