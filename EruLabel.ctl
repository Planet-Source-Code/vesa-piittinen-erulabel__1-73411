VERSION 5.00
Begin VB.UserControl EruLabel 
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  'None
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Windowless      =   -1  'True
End
Attribute VB_Name = "EruLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'* EruLabel - First Release [2010-09-04]                                                            *
'* ------------------------------------------------------------------------------------------------ *
'* By Vesa Piittinen aka Merri. Gmail account is vesa.piittinen                                     *
'*                                                                                                  *
'* Should work in Windows 2000 and later OS.                                                        *
'*                                                                                                  *
'* LICENSE                                                                                          *
'* ------------------------------------------------------------------------------------------------ *
'* http://creativecommons.org/licenses/by-sa/1.0/fi/deed.en                                         *
'*                                                                                                  *
'* Terms: 1) If you make your own version, share using this same license.                           *
'*        2) When used in a program, mention my name in the program's credits.                      *
'*        3) May not be used as a part of commercial controls suite.                                *
'*        4) Free for any other commercial and non-commercial usage.                                *
'*        5) Use at your own risk. No support guaranteed.                                           *
'*                                                                                                  *
'* REQUIREMENTS                                                                                     *
'* ------------------------------------------------------------------------------------------------ *
'* - No special requirements.                                                                       *
'*                                                                                                  *
'* HOW TO ADD TO YOUR PROGRAM                                                                       *
'* ------------------------------------------------------------------------------------------------ *
'* 1) Copy EruLabel.ctl and EruLabel.ctx to your project folder.                                    *
'* 2) In your project add EruLabel.ctl.                                                             *
'*                                                                                                  *
'* VERSION HISTORY                                                                                  *
'* ------------------------------------------------------------------------------------------------ *
'* 2010-09-04 First Release                                                                         *
'* - Alignment, AutoSize, BackColor, BackStyle, BorderColor, BorderRadius, BorderStyle, BorderWidth,*
'*   Caption, CaptionHex, CodePoints, Enabled, Font, ForeColor, Margin, MouseIcon, MousePointer,    *
'*   Opacity, Padding, RightToLeft, UseMnemonic, WordWrap                                           *
'* - Unicode aware                                                                                  *
'* - Issue: CodePoints gives UTF-16, but it should give UTF-32 (ie. true Unicode code points)       *
'* - Issue: does not support Data and Link properties (regular Label supports these)                *
'* - Issue: setting Enabled = False is not shown visually                                           *
'*                                                                                                  *
'* CREDITS                                                                                          *
'* ------------------------------------------------------------------------------------------------ *
'* - LaVolpe for his post & help on achieving windowless control transparency.                      *
'****************************************************************************************************
Option Explicit

Public Event Change()
Public Event Click()
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Public Event DblClick()
Attribute DblClick.VB_UserMemId = -601
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607

Public Enum LabelBackStyleConstants
    [Label Transparent] = 0
    [Label Opaque] = 1
End Enum

' API drawtext constants
Private Const DT_CALCRECT = &H400&
Private Const DT_CENTER = &H1&
Private Const DT_LEFT = &H0&
Private Const DT_NOCLIP = &H100&
Private Const DT_NOPREFIX = &H800&
Private Const DT_RIGHT = &H2&
Private Const DT_SINGLELINE = &H20
Private Const DT_WORDBREAK = &H10&

' API pen constants
Private Const PS_SOLID = 0
Private Const PS_DASH = 1
Private Const PS_DOT = 2
Private Const PS_DASHDOT = 3
Private Const PS_DASHDOTDOT = 4
Private Const PS_NULL = 5

Public Enum LabelBorderStyleConstants
    [Label Solid] = PS_SOLID
    [Label Dash] = PS_DASH
    [Label Dot] = PS_DOT
    [Label Dash Dot] = PS_DASHDOT
    [Label Dash Dot Dot] = PS_DASHDOTDOT
End Enum

' API font constants
Private Const FW_BOLD = 700&
Private Const FW_NORMAL = 400&
Private Const LF_FACESIZE = 32&
Private Const LOGPIXELSX = 88&
Private Const LOGPIXELSY = 90&

' API font structure
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
    lfFaceName(31) As Byte
End Type

' API RightToLeft view
Private Const TA_RTLREADING = &H100&

' API declaration
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hSrcDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As Any, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As Any, ByVal hBrush As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextAlign Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

' internal variables
Private m_BufferBitmap As Long
Private m_BufferDC As Long
Private m_BufferRECT(0 To 3) As Long
Private m_BufferResize As Boolean
Private m_ContainerScaleMode As ScaleModeConstants
Private m_FontHandle As Long
Private m_Format As Long

' properties
Private m_Alignment As AlignmentConstants
Private m_AutoSize As Boolean
Private m_BackColor As Long
Private m_BackStyle As Long
Private m_BorderRadius As Long
Private m_BorderColor As Long
Private m_BorderStyle As Long
Private m_BorderWidth As Long
Private m_Caption() As Byte
Private m_ForeColor As Long
Private WithEvents m_Font As StdFont
Attribute m_Font.VB_VarHelpID = -1
Private m_MarginBottom As Byte
Private m_MarginLeft As Byte
Private m_MarginRight As Byte
Private m_MarginTop As Byte
Private m_Opacity As Single
Private m_PaddingBottom As Byte
Private m_PaddingLeft As Byte
Private m_PaddingRight As Byte
Private m_PaddingTop As Byte
Private m_RightToLeft As Boolean
Private m_UseMnemonic As Boolean
Private m_WordWrap As Boolean

Public Property Get Alignment() As AlignmentConstants
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal NewValue As AlignmentConstants)
    m_Alignment = NewValue
    UpdateFormat
    UserControl.Refresh
End Property

Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_UserMemId = -500
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal NewValue As Boolean)
    m_AutoSize = NewValue
    UpdateFormat
    If m_AutoSize Then UpdateSize
    UserControl.Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    UserControl.BackColor = NewValue
    If NewValue < 0 Then m_BackColor = GetSysColor(NewValue And &HFF&) Else m_BackColor = NewValue
    UserControl.Refresh
End Property

Public Property Get BackStyle() As LabelBackStyleConstants
Attribute BackStyle.VB_UserMemId = -502
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal NewValue As LabelBackStyleConstants)
    m_BackStyle = NewValue
    UserControl.Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = UserControl.FillColor
End Property

Public Property Let BorderColor(ByVal NewValue As OLE_COLOR)
    UserControl.FillColor = NewValue
    If NewValue < 0 Then m_BorderColor = GetSysColor(NewValue And &HFF&) Else m_BorderColor = NewValue
    UserControl.Refresh
End Property

Public Property Get BorderRadius() As Byte
    BorderRadius = m_BorderRadius
End Property

Public Property Let BorderRadius(ByVal NewValue As Byte)
    m_BorderRadius = NewValue
    UserControl.Refresh
End Property

Public Property Get BorderStyle() As LabelBorderStyleConstants
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal NewValue As LabelBorderStyleConstants)
    m_BorderStyle = NewValue
    UserControl.Refresh
End Property

Public Property Get BorderWidth() As Byte
    BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal NewValue As Byte)
    m_BorderWidth = NewValue
    UserControl.Refresh
End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"
    Caption = m_Caption
End Property

Public Property Let Caption(NewValue As String)
    m_Caption = NewValue
    If m_UseMnemonic Then UpdateAccessKeys NewValue Else UpdateAccessKeys vbNullString
    If m_AutoSize Then UpdateSize
    UserControl.Refresh
    RaiseEvent Change
End Property

Public Property Get CaptionHex() As String
    Dim B As Byte, I As Long, Hex() As Byte
    If UBound(m_Caption) >= 1 Then
        ReDim Hex(UBound(m_Caption) * 4 + 3)
        For I = 0 To UBound(m_Caption)
            B = (m_Caption(I) And &HF0) \ &H10
            Select Case B
            Case 0 To 9
                Hex(I * 4) = B Or 48
            Case 10 To 15
                Hex(I * 4) = B + 55
            End Select
            B = m_Caption(I) And &HF
            Select Case B
            Case 0 To 9
                Hex(I * 4 + 2) = B Or 48
            Case 10 To 15
                Hex(I * 4 + 2) = B + 55
            End Select
        Next I
        CaptionHex = Hex
    End If
End Property

Public Property Let CaptionHex(NewValue As String)
    Dim B As Byte, I As Long, Hex() As Byte, N As Byte
    Hex = UCase$(NewValue)
    If UBound(Hex) >= 3 Then
        ReDim m_Caption(0 To (UBound(Hex) - 3) \ 4)
        For I = 0 To UBound(m_Caption)
            B = Hex(I * 4)
            Select Case B
            Case 48 To 57
                N = (B - 48) * &H10
            Case 65 To 70
                N = (B - 55) * &H10
            Case Else
                N = 0
            End Select
            B = Hex(I * 4 + 2)
            Select Case B
            Case 48 To 57
                N = N Or (B - 48)
            Case 65 To 70
                N = N Or (B - 55)
            End Select
            m_Caption(I) = N
        Next I
    Else
        m_Caption = vbNullString
    End If
    If m_UseMnemonic Then UpdateAccessKeys CStr(m_Caption) Else UpdateAccessKeys vbNullString
    If m_AutoSize Then UpdateSize
    UserControl.Refresh
    RaiseEvent Change
End Property

' needs to be fixed to UTF-32 (to give true Unicode code points)
Public Property Get CodePoints() As String
    Dim CP() As String, I As Long, UB As Long
    UB = UBound(m_Caption)
    If (UB > 0) And (UB And 1) = 1 Then
        ReDim CP((UB - 1) \ 2)
        For I = 0 To UBound(CP)
            CP(I) = (m_Caption(I * 2) And &HFF&) Or (m_Caption(I * 2 + 1) And &HFF&) * &H100&
        Next I
        CodePoints = Join(CP, ",")
    End If
End Property

' needs to be fixed to UTF-32 (to give true Unicode code points)
Public Property Let CodePoints(NewValue As String)
    Dim B As Byte, C As Long, CP() As String, I As Long
    If LenB(NewValue) = 0 Then
        m_Caption = vbNullString
    Else
        CP = Split(NewValue, ",")
        ReDim m_Caption(0 To UBound(CP) * 2 + 1)
        For I = 0 To UBound(CP)
            C = Val(CP(I))
            B = C And &HFF&
            m_Caption(I * 2) = B
            B = (C And &HFF00&) \ &H100&
            m_Caption(I * 2 + 1) = B
        Next I
    End If
    If m_UseMnemonic Then UpdateAccessKeys CStr(m_Caption) Else UpdateAccessKeys vbNullString
    If m_AutoSize Then UpdateSize
    UserControl.Refresh
    RaiseEvent Change
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    UserControl.Enabled = NewValue
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(NewValue As StdFont)
    Dim NewFont As New StdFont
    ' have to do it this way because otherwise we'd link with existing font object
    NewFont.Bold = NewValue.Bold
    NewFont.Charset = NewValue.Charset
    NewFont.Italic = NewValue.Italic
    NewFont.Name = NewValue.Name
    NewFont.Size = NewValue.Size
    NewFont.Strikethrough = NewValue.Strikethrough
    NewFont.Underline = NewValue.Underline
    NewFont.Weight = NewValue.Weight
    Set m_Font = NewFont
    m_Font_FontChanged vbNullString
    If m_AutoSize Then UpdateSize
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    UserControl.ForeColor = NewValue
    If NewValue < 0 Then m_ForeColor = GetSysColor(NewValue And &HFF&) Else m_ForeColor = NewValue
    UserControl.Refresh
End Property

Public Property Get Margin() As String
    Margin = m_MarginTop & "px " & m_MarginRight & "px " & m_MarginBottom & "px " & m_MarginLeft & "px"
End Property

Public Property Let Margin(ByVal NewValue As String)
    Dim I As Long, M() As String, U As Long
    NewValue = LCase$(Trim$(NewValue))
    If LenB(NewValue) = 0 Then
        m_MarginBottom = 0
        m_MarginLeft = 0
        m_MarginRight = 0
        m_MarginTop = 0
    Else
        Do While InStr(NewValue, "  "): NewValue = Replace(NewValue, "  ", " "): Loop
        M = Split(NewValue, "px")
        U = UBound(M)
        If U <> 3 Then
            ReDim Preserve M(3)
            For I = 0 To 2 - U: M(U + I + 1) = M(I): Next I
        End If
        m_MarginTop = Val(M(0))
        m_MarginRight = Val(M(1))
        m_MarginBottom = Val(M(2))
        m_MarginLeft = Val(M(3))
    End If
    If m_AutoSize Then UpdateSize
    UserControl.Refresh
End Property


Public Property Get MarginBottom() As Byte
Attribute MarginBottom.VB_MemberFlags = "400"
    MarginBottom = m_MarginBottom
End Property

Public Property Let MarginBottom(ByVal NewValue As Byte)
    m_MarginBottom = NewValue
    If m_AutoSize Then UpdateSize
    UserControl.Refresh
End Property

Public Property Get MarginLeft() As Byte
Attribute MarginLeft.VB_MemberFlags = "400"
    MarginLeft = m_MarginLeft
End Property

Public Property Let MarginLeft(ByVal NewValue As Byte)
    m_MarginLeft = NewValue
    If m_AutoSize Then UpdateSize
    UserControl.Refresh
End Property

Public Property Get MarginRight() As Byte
Attribute MarginRight.VB_MemberFlags = "400"
    MarginRight = m_MarginRight
End Property

Public Property Let MarginRight(ByVal NewValue As Byte)
    m_MarginRight = NewValue
    If m_AutoSize Then UpdateSize
    UserControl.Refresh
End Property

Public Property Get MarginTop() As Byte
Attribute MarginTop.VB_MemberFlags = "400"
    MarginTop = m_MarginTop
End Property

Public Property Let MarginTop(ByVal NewValue As Byte)
    m_MarginTop = NewValue
    If m_AutoSize Then UpdateSize
    UserControl.Refresh
End Property

Public Property Get MouseIcon() As IPictureDisp
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(NewValue As IPictureDisp)
    Set UserControl.MouseIcon = NewValue
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal NewValue As MousePointerConstants)
    UserControl.MousePointer = NewValue
End Property

Public Property Get Opacity() As Single
    Opacity = m_Opacity
End Property

Public Property Let Opacity(ByVal NewValue As Single)
    If NewValue < 0 Then
        NewValue = 0
    ElseIf NewValue > 1 Then
        NewValue = 1
    End If
    m_Opacity = NewValue
    UserControl.Refresh
End Property

Public Property Get Padding() As String
    Padding = m_PaddingTop & "px " & m_PaddingRight & "px " & m_PaddingBottom & "px " & m_PaddingLeft & "px"
End Property

Public Property Let Padding(ByVal NewValue As String)
    Dim I As Long, M() As String, U As Long
    NewValue = LCase$(Trim$(NewValue))
    If LenB(NewValue) = 0 Then
        m_PaddingBottom = 0
        m_PaddingLeft = 0
        m_PaddingRight = 0
        m_PaddingTop = 0
    Else
        Do While InStr(NewValue, "  "): NewValue = Replace(NewValue, "  ", " "): Loop
        M = Split(NewValue, "px")
        U = UBound(M)
        If U <> 3 Then
            ReDim Preserve M(3)
            For I = 0 To 2 - U: M(U + I + 1) = M(I): Next I
        End If
        m_PaddingTop = Val(M(0))
        m_PaddingRight = Val(M(1))
        m_PaddingBottom = Val(M(2))
        m_PaddingLeft = Val(M(3))
    End If
    If m_AutoSize Then UpdateSize
    UserControl.Refresh
End Property

Public Property Get PaddingBottom() As Byte
Attribute PaddingBottom.VB_MemberFlags = "400"
    PaddingBottom = m_PaddingBottom
End Property

Public Property Let PaddingBottom(ByVal NewValue As Byte)
    m_PaddingBottom = NewValue
    If m_AutoSize Then UpdateSize
    UserControl.Refresh
End Property

Public Property Get PaddingLeft() As Byte
Attribute PaddingLeft.VB_MemberFlags = "400"
    PaddingLeft = m_PaddingLeft
End Property

Public Property Let PaddingLeft(ByVal NewValue As Byte)
    m_PaddingLeft = NewValue
    If m_AutoSize Then UpdateSize
    UserControl.Refresh
End Property

Public Property Get PaddingRight() As Byte
Attribute PaddingRight.VB_MemberFlags = "400"
    PaddingRight = m_PaddingRight
End Property

Public Property Let PaddingRight(ByVal NewValue As Byte)
    m_PaddingRight = NewValue
    If m_AutoSize Then UpdateSize
    UserControl.Refresh
End Property

Public Property Get PaddingTop() As Byte
Attribute PaddingTop.VB_MemberFlags = "400"
    PaddingTop = m_PaddingTop
End Property

Public Property Let PaddingTop(ByVal NewValue As Byte)
    m_PaddingTop = NewValue
    If m_AutoSize Then UpdateSize
    UserControl.Refresh
End Property

Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_UserMemId = -611
    RightToLeft = m_RightToLeft
End Property

Public Property Let RightToLeft(ByVal NewValue As Boolean)
    m_RightToLeft = NewValue
    If m_BufferDC Then
        ' update RTL information for the buffer
        If m_RightToLeft Then
            SetTextAlign m_BufferDC, (GetTextAlign(m_BufferDC) Or TA_RTLREADING)
        Else
            SetTextAlign m_BufferDC, (GetTextAlign(m_BufferDC) And Not TA_RTLREADING)
        End If
        If m_AutoSize Then UpdateSize
    End If
    UserControl.Refresh
End Property

Private Sub UpdateAccessKeys(NewValue As String)
    Dim I As Long, NewKeys As String, S As String
    If StrPtr(NewValue) Then
        I = InStr(NewValue, "&")
        Do While I < Len(NewValue) And I > 0
            I = I + 1
            S = Mid$(NewValue, I, 1)
            If AscW(S) = 38 Then
                I = InStr(I + 1, NewValue, "&")
            Else
                NewKeys = UCase$(S)
                Exit Do
            End If
        Loop
    End If
    UserControl.AccessKeys = NewKeys
End Sub

Private Sub UpdateFont()
    Dim LogicalFont As LOGFONT, lngLen As Long
    ' initialize font settings
    With m_Font
        ' determine length of font name
        If Len(.Name) >= LF_FACESIZE Then lngLen = LF_FACESIZE Else lngLen = Len(.Name)
        ' copy maximum allowed length
        CopyMemory LogicalFont.lfFaceName(0), ByVal .Name, lngLen
        ' set other font settings
        LogicalFont.lfHeight = Int(UserControl.ScaleY(-.Size / 72, vbInches, vbPixels))
        LogicalFont.lfItalic = .Italic
        LogicalFont.lfWeight = .Weight
        LogicalFont.lfUnderline = .Underline
        LogicalFont.lfStrikeOut = .Strikethrough
        LogicalFont.lfCharSet = .Charset
        LogicalFont.lfQuality = 6
    End With
    ' destroy old font if it exists
    If m_FontHandle Then DeleteObject m_FontHandle
    ' create new font
    m_FontHandle = CreateFontIndirect(LogicalFont)
End Sub

Private Sub UpdateFormat()
    ' initialize with alignment
    If m_Alignment = vbLeftJustify Then
        m_Format = DT_LEFT
    ElseIf m_Alignment = vbCenter Then
        m_Format = DT_CENTER
    Else
        m_Format = DT_RIGHT
    End If
    ' & character for AccessKeys or not?
    If Not m_UseMnemonic Then m_Format = m_Format Or DT_NOPREFIX
    ' word wrapping?
    If m_WordWrap Then
        ' word wrap
        m_Format = m_Format Or DT_WORDBREAK
    ElseIf m_AutoSize Then
        ' force to size if autosizing and not wrapping
        m_Format = m_Format Or DT_NOCLIP
    End If
End Sub

Private Sub UpdateSize()
    Dim NewHeight As Single, NewWidth As Single, OldFont As Long, TextRect(0 To 3) As Long
    If m_AutoSize And (m_BufferDC <> 0) Then
        ' calculate width & height of the text
        TextRect(2) = UserControl.ScaleWidth - m_MarginRight - m_PaddingRight - m_MarginLeft - m_PaddingLeft - m_BorderWidth * 2
        'TextRect(3) = m_BufferRECT(3) - m_PaddingBottom - m_MarginTop - m_PaddingTop - m_BorderWidth * 2
        ' change font
        OldFont = SelectObject(m_BufferDC, m_FontHandle)
        ' calculate size for font rectangle
        DrawTextW m_BufferDC, VarPtr(m_Caption(0)), (UBound(m_Caption) + 1) \ 2, TextRect(0), m_Format Or DT_CALCRECT
        ' restore original font
        SelectObject m_BufferDC, OldFont
        ' calculate new height
        m_BufferRECT(3) = TextRect(3) + m_MarginBottom + m_PaddingBottom + m_MarginTop + m_PaddingTop + m_BorderWidth * 2
        NewHeight = ScaleY(m_BufferRECT(3), vbPixels, vbTwips)
        ' mark resize true
        m_BufferResize = True
        ' then see if we resize height only or both dimensions
        If m_WordWrap Then
            m_BufferRECT(2) = UserControl.ScaleWidth
            ' resize height only, resize only if we need to do so
            If NewHeight <> UserControl.Height Then UserControl.Height = NewHeight
        Else
            ' resize both width & height, calculate new width
            m_BufferRECT(2) = TextRect(2) + m_MarginRight + m_PaddingRight + m_MarginLeft + m_PaddingLeft + m_BorderWidth * 2
            NewWidth = ScaleX(m_BufferRECT(2), vbPixels, vbTwips)
            ' resize only if we need to do so
            If (NewWidth <> UserControl.Width) Or (NewHeight <> UserControl.Height) Then UserControl.Size NewWidth, NewHeight
        End If
    Else
        ' set drawing boundaries
        m_BufferRECT(2) = UserControl.ScaleWidth
        m_BufferRECT(3) = UserControl.ScaleHeight
    End If
End Sub

Public Property Get UseMnemonic() As Boolean
    UseMnemonic = m_UseMnemonic
End Property

Public Property Let UseMnemonic(ByVal NewValue As Boolean)
    m_UseMnemonic = NewValue
    If m_UseMnemonic Then UpdateAccessKeys CStr(m_Caption) Else UpdateAccessKeys vbNullString
    UpdateFormat
    If m_AutoSize Then UpdateSize
    UserControl.Refresh
End Property

Public Property Get WordWrap() As Boolean
    WordWrap = m_WordWrap
End Property

Public Property Let WordWrap(ByVal NewValue As Boolean)
    m_WordWrap = NewValue
    UpdateFormat
    If m_AutoSize Then UpdateSize
    UserControl.Refresh
End Property

Private Sub m_Font_FontChanged(ByVal PropertyName As String)
    UpdateFont
    ' update control using the new font
    UserControl.Refresh
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    ' refresh color information
    Select Case PropertyName
    Case "ScaleUnits"
        Select Case UserControl.Ambient.ScaleUnits
        Case "Twip"
            m_ContainerScaleMode = vbTwips
        Case "Point"
            m_ContainerScaleMode = vbPoints
        Case "Pixel"
            m_ContainerScaleMode = vbPixels
        Case "Character"
            m_ContainerScaleMode = vbCharacters
        Case "Inch"
            m_ContainerScaleMode = vbInches
        Case "Millimeter"
            m_ContainerScaleMode = vbMillimeters
        Case "Centimeter"
            m_ContainerScaleMode = vbCentimeters
        Case "User"
            m_ContainerScaleMode = vbUser
        End Select
    Case Else
        BackColor = UserControl.BackColor
        BorderColor = UserControl.FillColor
        ForeColor = UserControl.ForeColor
    End Select
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    HitResult = vbHitResultHit
End Sub

Private Sub UserControl_Initialize()
    ' create empty byte array to avoid errors
    m_Caption = vbNullString
    ' default to twips
    m_ContainerScaleMode = vbTwips
End Sub

Private Sub UserControl_InitProperties()
    ' update m_ContainerScaleMode
    UserControl_AmbientChanged "ScaleUnits"
    ' modernization effort
    With UserControl.Extender.Container
        ' replace container's font with Segoe UI!
        If .Font.Name = "MS Sans Serif" And .Font.Size = 8 Then
            ' EruLabel's default font is Segoe UI using size 9
            Set .Font = UserControl.Font
        End If
    End With
    ' initial back style
    m_BackStyle = [Label Opaque]
    ' we must see something!
    m_Opacity = 1
    ' use Mnemonic
    m_UseMnemonic = True
    ' initial name
    m_Caption = UserControl.Ambient.DisplayName
    ' update colors
    BackColor = UserControl.Ambient.BackColor
    BorderColor = UserControl.FillColor
    ForeColor = UserControl.Ambient.ForeColor
    ' change font
    Set m_Font = UserControl.Ambient.Font
    ' update these
    UpdateFont
    UpdateFormat
    UpdateSize
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NewX As Single, NewY As Single
    If m_ContainerScaleMode <> vbUser Then
        NewX = UserControl.ScaleX(X, vbPixels, m_ContainerScaleMode)
        NewY = UserControl.ScaleY(Y, vbPixels, m_ContainerScaleMode)
    Else
        On Error Resume Next
        NewX = UserControl.Extender.Container.ScaleX(X, vbPixels, m_ContainerScaleMode)
        NewY = UserControl.Extender.Container.ScaleY(Y, vbPixels, m_ContainerScaleMode)
        On Error GoTo 0
    End If
    RaiseEvent MouseDown(Button, Shift, NewX, NewY)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NewX As Single, NewY As Single
    If m_ContainerScaleMode <> vbUser Then
        NewX = UserControl.ScaleX(X, vbPixels, m_ContainerScaleMode)
        NewY = UserControl.ScaleY(Y, vbPixels, m_ContainerScaleMode)
    Else
        On Error Resume Next
        NewX = UserControl.Extender.Container.ScaleX(X, vbPixels, m_ContainerScaleMode)
        NewY = UserControl.Extender.Container.ScaleY(Y, vbPixels, m_ContainerScaleMode)
        On Error GoTo 0
    End If
    RaiseEvent MouseMove(Button, Shift, NewX, NewY)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NewX As Single, NewY As Single
    If m_ContainerScaleMode <> vbUser Then
        NewX = UserControl.ScaleX(X, vbPixels, m_ContainerScaleMode)
        NewY = UserControl.ScaleY(Y, vbPixels, m_ContainerScaleMode)
    Else
        On Error Resume Next
        NewX = UserControl.Extender.Container.ScaleX(X, vbPixels, m_ContainerScaleMode)
        NewY = UserControl.Extender.Container.ScaleY(Y, vbPixels, m_ContainerScaleMode)
        On Error GoTo 0
    End If
    RaiseEvent MouseUp(Button, Shift, NewX, NewY)
End Sub

Private Sub UserControl_Paint()
    Dim BackColorBrush As Long, BorderPen As Long
    Dim DC As Long, OldFont As Long
    Dim WithBorder(0 To 3) As Long, WithMargin(0 To 3) As Long, WithPadding(0 To 3) As Long
    ' cache DC
    DC = UserControl.hDC
    ' font, the precious font...
    If m_FontHandle = 0 Then UpdateFont
    ' do we have any text to draw?
    If UBound(m_Caption) >= 0 And m_Opacity > 0 Then
        ' margin
        WithMargin(0) = m_MarginLeft
        WithMargin(1) = m_MarginTop
        WithMargin(2) = m_BufferRECT(2) - m_MarginRight
        WithMargin(3) = m_BufferRECT(3) - m_MarginBottom
        ' padding
        WithPadding(0) = m_MarginLeft + m_PaddingLeft + m_BorderWidth
        WithPadding(1) = m_MarginTop + m_PaddingTop + m_BorderWidth
        WithPadding(2) = WithMargin(2) - m_PaddingRight - m_BorderWidth
        WithPadding(3) = WithMargin(3) - m_PaddingBottom - m_BorderWidth
        ' create a buffer DC & bitmap, pair them together
        If m_BufferDC = 0 Then
            m_BufferDC = CreateCompatibleDC(DC)
            ' set text RTL
            If m_RightToLeft Then
                SetTextAlign m_BufferDC, (GetTextAlign(m_BufferDC) Or TA_RTLREADING)
            Else
                SetTextAlign m_BufferDC, (GetTextAlign(m_BufferDC) And Not TA_RTLREADING)
            End If
            ' first size update!
            UpdateSize
        End If
        ' do we have a compatible bitmap? do we need a bitmap resize?
        If m_BufferBitmap = 0 Or m_BufferResize Then
            m_BufferResize = False
            m_BufferBitmap = CreateCompatibleBitmap(DC, m_BufferRECT(2), m_BufferRECT(3))
            DeleteObject SelectObject(m_BufferDC, m_BufferBitmap)
        End If
        ' copy initial state
        BitBlt m_BufferDC, 0, 0, m_BufferRECT(2), m_BufferRECT(3), DC, 0, 0, vbSrcCopy
        ' set drawing mode to transparent
        SetBkMode m_BufferDC, 3
        SetBkColor m_BufferDC, m_BackColor
        ' do we draw the background & border?
        If m_BackStyle = [Label Opaque] Then
            ' are we drawing rounded?
            If m_BorderRadius = 0 Then
                If m_BorderWidth = 0 Then
                    ' create a brush for FillRect
                    BackColorBrush = CreateSolidBrush(m_BackColor)
                    ' draw the background
                    FillRect m_BufferDC, WithMargin(0), BackColorBrush
                    ' remove the brush
                    DeleteObject BackColorBrush
                Else
                    ' create a background color brush
                    DeleteObject SelectObject(m_BufferDC, CreateSolidBrush(m_BackColor))
                    ' create a pen for border
                    If m_BorderWidth = 0 Then
                        BorderPen = CreatePen(PS_NULL, 0, m_BorderColor)
                    Else
                        BorderPen = CreatePen(m_BorderStyle, m_BorderWidth * 2 - 1, m_BorderColor)
                    End If
                    ' apply the pen, remove the old
                    DeleteObject SelectObject(m_BufferDC, BorderPen)
                    ' draw a rectangle
                    Rectangle m_BufferDC, WithMargin(0), WithMargin(1), WithMargin(2), WithMargin(3)
                End If
            Else
                ' create a background color brush
                DeleteObject SelectObject(m_BufferDC, CreateSolidBrush(m_BackColor))
                ' create a pen for border
                If m_BorderWidth = 0 Then
                    BorderPen = CreatePen(PS_NULL, 0, m_BorderColor)
                Else
                    BorderPen = CreatePen(m_BorderStyle, m_BorderWidth * 2 - 1, m_BorderColor)
                End If
                ' apply the pen, remove the old
                DeleteObject SelectObject(m_BufferDC, BorderPen)
                ' draw a rounded rectangle
                RoundRect m_BufferDC, WithMargin(0), WithMargin(1), WithMargin(2), WithMargin(3), m_BorderRadius * 2.5, m_BorderRadius * 2.5
            End If
        End If
        ' set font color
        SetTextColor m_BufferDC, m_ForeColor
        ' replace font with our own
        OldFont = SelectObject(m_BufferDC, m_FontHandle)
        ' draw the caption text
        DrawTextW m_BufferDC, VarPtr(m_Caption(0)), (UBound(m_Caption) + 1) \ 2, WithPadding(0), m_Format
        ' restore original font (which is then destroyed as the DC we just used is destroyed)
        SelectObject m_BufferDC, OldFont
        ' copy modified image back
        If m_Opacity < 1 Then
            AlphaBlend DC, 0, 0, m_BufferRECT(2), m_BufferRECT(3), m_BufferDC, 0, 0, m_BufferRECT(2), m_BufferRECT(3), CLng(m_Opacity * 255) * &H10000
        Else
            BitBlt DC, 0, 0, m_BufferRECT(2), m_BufferRECT(3), m_BufferDC, 0, 0, vbSrcCopy
        End If
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' update m_ContainerScaleMode
    UserControl_AmbientChanged "ScaleUnits"
    ' then get our stuff!
    With PropBag
        m_Alignment = .ReadProperty("Alignment", vbLeftJustify)
        m_AutoSize = .ReadProperty("AutoSize", False)
        BackColor = .ReadProperty("BackColor", vbButtonFace)
        m_BackStyle = .ReadProperty("BackStyle", 0)
        BorderColor = .ReadProperty("BorderColor", 0)
        m_BorderRadius = .ReadProperty("BorderRadius", 0)
        m_BorderStyle = .ReadProperty("BorderStyle", 0)
        m_BorderWidth = .ReadProperty("Borderwidth", 0)
        m_Caption = .ReadProperty("Caption", UserControl.Ambient.DisplayName)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        Set m_Font = .ReadProperty("Font", UserControl.Ambient.Font)
        ForeColor = .ReadProperty("ForeColor", vbButtonText)
        m_MarginBottom = .ReadProperty("MarginBottom", 0)
        m_MarginLeft = .ReadProperty("MarginLeft", 0)
        m_MarginRight = .ReadProperty("MarginRight", 0)
        m_MarginTop = .ReadProperty("MarginTop", 0)
        Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
        m_Opacity = .ReadProperty("Opacity", 1)
        m_PaddingBottom = .ReadProperty("PaddingBottom", 0)
        m_PaddingLeft = .ReadProperty("PaddingLeft", 0)
        m_PaddingRight = .ReadProperty("PaddingRight", 0)
        m_PaddingTop = .ReadProperty("PaddingTop", 0)
        m_RightToLeft = .ReadProperty("RightToLeft", False)
        m_UseMnemonic = .ReadProperty("UseMnemonic", True)
        m_WordWrap = .ReadProperty("WordWrap", False)
    End With
    ' update these
    UpdateFont
    UpdateFormat
    UpdateSize
End Sub

Private Sub UserControl_Resize()
    ' mark resize true
     m_BufferResize = True
    ' update size information
    UpdateSize
End Sub

Private Sub UserControl_Terminate()
    ' destroy the font we have created
    If m_FontHandle Then DeleteObject m_FontHandle
    ' delete buffer
    If m_BufferDC Then DeleteDC m_BufferDC
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ' save our stuff!
    With PropBag
        .WriteProperty "Alignment", m_Alignment, vbLeftJustify
        .WriteProperty "AutoSize", m_AutoSize, False
        .WriteProperty "BackColor", UserControl.BackColor, vbButtonFace
        .WriteProperty "BackStyle", m_BackStyle, 0
        .WriteProperty "BorderColor", UserControl.FillColor, 0
        .WriteProperty "BorderRadius", m_BorderRadius, 0
        .WriteProperty "BorderStyle", m_BorderStyle, 0
        .WriteProperty "BorderWidth", m_BorderWidth, 0
        .WriteProperty "Caption", m_Caption, vbNullString
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "Font", m_Font, UserControl.Ambient.Font
        .WriteProperty "ForeColor", UserControl.ForeColor, vbButtonText
        .WriteProperty "MarginBottom", m_MarginBottom, 0
        .WriteProperty "MarginLeft", m_MarginLeft, 0
        .WriteProperty "MarginRight", m_MarginRight, 0
        .WriteProperty "MarginTop", m_MarginTop, 0
        .WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
        .WriteProperty "MousePointer", UserControl.MousePointer, vbDefault
        .WriteProperty "Opacity", m_Opacity, 1
        .WriteProperty "PaddingBottom", m_PaddingBottom, 0
        .WriteProperty "PaddingLeft", m_PaddingLeft, 0
        .WriteProperty "PaddingRight", m_PaddingRight, 0
        .WriteProperty "PaddingTop", m_PaddingTop, 0
        .WriteProperty "RightToLeft", m_RightToLeft, False
        .WriteProperty "UseMnemonic", m_UseMnemonic, True
        .WriteProperty "WordWrap", m_WordWrap, False
    End With
End Sub
