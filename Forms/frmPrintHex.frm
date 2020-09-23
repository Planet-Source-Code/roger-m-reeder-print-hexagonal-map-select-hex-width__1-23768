VERSION 5.00
Begin VB.Form frmPrintHex 
   Caption         =   "Print Hex Map"
   ClientHeight    =   885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   885
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Line Width "
      Height          =   615
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   1095
      Begin VB.TextBox txtDrawWidth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "1"
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Hex Width "
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.ComboBox cboScaleMode 
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtHexWidth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "0.25"
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPrintHex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()
    Dim Xmax As Long
    Dim Ymax As Long
    Dim XCount As Long
    Dim YCount As Long
    On Error GoTo Error_cmdPrint
    
    Dim X1 As Long
    Dim Y1 As Long
    Dim X2 As Long
    Dim Y2 As Long
    Dim X As Long
    Dim U As Long   'Half of Hex width horizontally
    Dim V As Long   'Half of hex width vertically
    Dim r As Long   'horizontal distance from left most point of hex to top left edge of hex
    ' --- |-R-/---U---\             /---
    '  |  |  /         \           /
    '  V  | /           \         /
    ' _|_ |/             \_______/
    '      \             /       \
    '       \           /         \
    '        \         /           \
    '         \-------/             \---
    
    Dim HexWidth As Double
    Me.Caption = "Print Hex Map Printing..."
    Me.MousePointer = vbHourglass
    HexWidth = Val(Me.txtHexWidth)
    
    'Convert Hexwidth taking into account Printer Scale
    V = Printer.ScaleY(HexWidth, Me.cboScaleMode.ItemData(Me.cboScaleMode.ListIndex), Printer.ScaleMode) / 2
    r = Printer.ScaleX(HexWidth, Me.cboScaleMode.ItemData(Me.cboScaleMode.ListIndex), Printer.ScaleMode) / 2
    
    
    U = Tan(DegToRad(30)) * r
    
    Xmax = Printer.ScaleWidth / (2 * (U + r))
    Ymax = Printer.ScaleHeight / (V * 2)
    
    Printer.DrawWidth = Val(Me.txtDrawWidth)
    For XCount = 0 To Xmax
        For YCount = 0 To Ymax
            Printer.Line (XCount * 2 * (U + r) + U, YCount * V * 2 + 0)-(XCount * 2 * (U + r) + U + r, YCount * V * 2 + 0)
            Printer.Line -(XCount * 2 * (U + r) + 2 * U + r, YCount * V * 2 + V)
            Printer.Line -(XCount * 2 * (U + r) + U + r, YCount * V * 2 + 2 * V)
            Printer.Line -(XCount * 2 * (U + r) + U, YCount * V * 2 + 2 * V)
            Printer.Line -(XCount * 2 * (U + r) + 0, YCount * V * 2 + V)
            Printer.Line -(XCount * 2 * (U + r) + U, YCount * V * 2 + 0)
            Printer.Line (XCount * 2 * (U + r) + 2 * U + r, YCount * V * 2 + V)-(XCount * 2 * (U + r) + 2 * (U + r), YCount * V * 2 + V)
        Next YCount
    Next XCount
    Printer.EndDoc
    Me.Caption = "Print Hex Map"
    Me.MousePointer = vbNormal
    Exit Sub
    
Error_cmdPrint:
    Printer.KillDoc
    Printer.EndDoc
End Sub

Private Sub Form_Load()
    Dim cbo As ComboBox
    
    Set cbo = Me.cboScaleMode
    cbo.Clear
    cbo.AddItem "Inches"
    cbo.ItemData(cbo.NewIndex) = vbInches
    cbo.AddItem "Pixels"
    cbo.ItemData(cbo.NewIndex) = vbPixels
    cbo.AddItem "Centimeters"
    cbo.ItemData(cbo.NewIndex) = vbCentimeters
    cbo.AddItem "Millimeters"
    cbo.ItemData(cbo.NewIndex) = vbMillimeters
    cbo.AddItem "Points"
    cbo.ItemData(cbo.NewIndex) = vbPoints
    cbo.AddItem "Twips"
    cbo.ItemData(cbo.NewIndex) = vbTwips
    Me.cboScaleMode = "Inches"
    Me.cboScaleMode.ListIndex = 0
    Debug.Print Me.cboScaleMode.ItemData(Me.cboScaleMode.ListIndex)
End Sub
