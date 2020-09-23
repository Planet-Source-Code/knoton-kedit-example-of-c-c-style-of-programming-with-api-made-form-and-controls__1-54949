Attribute VB_Name = "modShowFont"
Option Explicit

Const DEFAULT_CHARSET = 1
Const OUT_DEFAULT_PRECIS = 0
Const CLIP_DEFAULT_PRECIS = 0
Const DEFAULT_QUALITY = 0
Const DEFAULT_PITCH = 0
Const FF_ROMAN = 16
Const CF_PRINTERFONTS = &H2
Const CF_SCREENFONTS = &H1
Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Const CF_EFFECTS = &H100&
Const CF_FORCEFONTEXIST = &H10000
Const CF_INITTOLOGFONTSTRUCT = &H40&
Const CF_LIMITSIZE = &H2000&
Const REGULAR_FONTTYPE = &H400
Const LF_FACESIZE = 32
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Const LOGPIXELSY = 90

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 31
End Type

Private Type CHOOSEFONT
        lStructSize As Long
        hwndOwner As Long          '  caller's window handle
        hDC As Long                '  printer DC/IC or NULL
        lpLogFont As Long          '  ptr. to a LOGFONT struct
        iPointSize As Long         '  10 * size in points of selected font
        flags As Long              '  enum. type flags
        rgbColors As Long          '  returned text color
        lCustData As Long          '  data passed to hook fn.
        lpfnHook As Long           '  ptr. to hook function
        lpTemplateName As String   '  custom template name
        hInstance As Long          '  instance handle of.EXE that
                                   '    contains cust. dlg. template
        lpszStyle As String        '  return the style field here
                                   '  must be LF_FACESIZE or bigger
        nFontType As Integer       '  same value reported to the EnumFonts
                                   '  call back with the extra FONTTYPE_
                                   '  bits added
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long           '  minimum pt size allowed &
        nSizeMax As Long           '  max pt size allowed if
                                       '    CF_LIMITSIZE is used
End Type

Public Enum FontStyle
    FW_DONTCARE = 0
    FW_THIN = 100
    FW_EXTRALIGHT = 200
    FW_LIGHT = 300
    FW_NORMAL = 400
    FW_MEDIUM = 500
    FW_SEMIBOLD = 600
    FW_BOLD = 700
    FW_EXTRABOLD = 800
    FW_HEAVY = 900
    FW_BLACK = FW_HEAVY
    FW_DEMIBOLD = FW_SEMIBOLD
    FW_REGULAR = FW_NORMAL
    FW_ULTRABOLD = FW_EXTRABOLD
    FW_ULTRALIGHT = FW_EXTRALIGHT
End Enum

Public Type FontType
    sName As String
    Size As Long
    mfont As FontStyle
    bItalic As Boolean
    bUnderline As Boolean
End Type

Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Public mFontType As FontType

Public Sub ShowFont(mFontStyle As FontStyle, bItalic As Boolean, bUnderlined As Boolean)
Dim cf As CHOOSEFONT, lfont As LOGFONT, hMem As Long, pMem As Long
Dim retval As Long, hFont As Long

lfont.lfHeight = -MulDiv((mFontType.Size), (GetDeviceCaps(GetDC(0), LOGPIXELSY)), 72)  ' Set height
lfont.lfWidth = 0  ' determine default width
lfont.lfEscapement = 0  ' angle between baseline and escapement vector
lfont.lfOrientation = 0  ' angle between baseline and orientation vector
lfont.lfWeight = mFontType.mfont  ' normal weight i.e. not bold
lfont.lfCharSet = DEFAULT_CHARSET  ' use default character set
lfont.lfOutPrecision = OUT_DEFAULT_PRECIS  ' default precision mapping
lfont.lfClipPrecision = CLIP_DEFAULT_PRECIS  ' default clipping precision
lfont.lfQuality = DEFAULT_QUALITY  ' default quality setting
lfont.lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN  ' default pitch, proportional with serifs
lfont.lfFaceName = mFontType.sName & vbNullChar   ' string must be null-terminated
lfont.lfItalic = Abs(mFontType.bItalic)
lfont.lfUnderline = Abs(mFontType.bUnderline)

' Create the memory block which will act as the LOGFONT structure buffer.
hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lfont))
pMem = GlobalLock(hMem)  ' lock and get pointer
CopyMemory ByVal pMem, lfont, Len(lfont)  ' copy structure's contents into block
' Initialize dialog box: Screen and printer fonts, point size between 10 and 72.
cf.lStructSize = Len(cf)  ' size of structure
cf.hwndOwner = m_hWnd  ' Our main window
cf.hDC = GetDC(0)  ' Set device context
cf.lpLogFont = pMem   ' pointer to LOGFONT memory block buffer
cf.iPointSize = 120  ' 12 point font (in units of 1/10 point)
cf.flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
cf.rgbColors = RGB(0, 0, 0)  ' black
cf.nFontType = REGULAR_FONTTYPE  ' regular font type i.e. not bold or anything
cf.nSizeMin = 10  ' minimum point size
cf.nSizeMax = 72  ' maximum point size

' Now, call the function.  If successful, copy the LOGFONT structure back into the structure
' and then print out the attributes we mentioned earlier that the user selected.
retval = CHOOSEFONT(cf)  ' open the dialog box
If retval <> 0 Then  ' success
    CopyMemory lfont, ByVal pMem, Len(lfont)  ' copy memory back
    
    With mFontType
        .bItalic = CBool(lfont.lfItalic)
        .bUnderline = CBool(lfont.lfUnderline)
        .mfont = lfont.lfWeight
        .sName = Left(lfont.lfFaceName, InStr(lfont.lfFaceName, vbNullChar) - 1)
        .Size = (cf.iPointSize / 10)
        hFont = CreateFont(lfont.lfHeight, 0, 0, 0, .mfont, lfont.lfItalic, lfont.lfUnderline, 0, 0, 0, 0, 0, 0, .sName)
        'set the font to the richedit
        SendMessage Edit_hWind, WM_SETFONT, hFont, 1
        SetFocus Edit_hWind
    End With
    
 End If
' Deallocate the memory block we created earlier.  Note that this must
' be done whether the function succeeded or not.
retval = GlobalUnlock(hMem)  ' destroy pointer, unlock block
retval = GlobalFree(hMem)  ' free the allocated memory
End Sub

