VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Rotate Text"
   ClientHeight    =   5175
   ClientLeft      =   3030
   ClientTop       =   1695
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5175
   ScaleWidth      =   6450
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Height          =   495
      Left            =   1740
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtSize 
      Height          =   285
      Left            =   840
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "12"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtDegree 
      Height          =   285
      Left            =   840
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "90"
      Top             =   60
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Degrees"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
' lfFaceName(LF_FACESIZE) As Byte 'THIS WAS DEFINED IN API-CHANGES MY OWN
  lfFacename As String * 33
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


Private Sub FontStuff()
  On Error GoTo GetOut
  Me.Cls
  Dim F As LOGFONT, hPrevFont As Long, hFont As Long, FontName As String
  Dim FONTSIZE As Integer
  FONTSIZE = Val(txtSize.Text)

  F.lfEscapement = 10 * Val(txtDegree.Text) 'rotation angle, in tenths
  FontName = "Arial Black" + Chr$(0) 'null terminated
  F.lfFacename = FontName
  F.lfHeight = (FONTSIZE * -20) / Screen.TwipsPerPixelY
  hFont = CreateFontIndirect(F)
  hPrevFont = SelectObject(Me.hdc, hFont)
  CurrentX = 3930
  CurrentY = 3860
  Print "SParq"
  
'  Clean up, restore original font
  hFont = SelectObject(Me.hdc, hPrevFont)
  DeleteObject hFont
  
  Exit Sub
GetOut:
  Exit Sub

End Sub

Private Sub Command1_Click()
  FontStuff
End Sub


Private Sub txtDegree_Change()
   If Val(txtDegree) < 1 Then txtDegree = 1: Exit Sub
   If Val(txtDegree) > 360 Then txtDegree = 360: Exit Sub
   Command1_Click
End Sub

Private Sub txtsize_Change()
  If Not IsNumeric(txtSize.Text) Then txtSize.Text = "18"
End Sub


