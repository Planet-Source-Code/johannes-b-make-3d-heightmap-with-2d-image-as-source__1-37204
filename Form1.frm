VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Heightmap Beta 3.5 by Johannes B 2002"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10875
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "A"
      Height          =   255
      Left            =   10560
      TabIndex        =   60
      ToolTipText     =   "About heightmap"
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Advanced settings..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   59
      Top             =   4800
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog CM 
      Left            =   10440
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "0"
      Height          =   255
      Left            =   10320
      TabIndex        =   55
      ToolTipText     =   "Reset scroll to 0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Caption         =   "Draw style"
      Height          =   1695
      Left            =   8400
      TabIndex        =   52
      Top             =   5160
      Width           =   2415
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   58
         Text            =   "1"
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Spikes (slow)"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Connected lines (slow)"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Dots"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "line/spike width"
         Height          =   255
         Left            =   600
         TabIndex        =   57
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Export heightmap to bitmap..."
      Height          =   375
      Left            =   7560
      TabIndex        =   35
      ToolTipText     =   "Save heightmap as bitmap image"
      Top             =   4320
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "C"
      Height          =   255
      Left            =   6960
      TabIndex        =   47
      ToolTipText     =   "Center view"
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Change heightmap field size..."
      Height          =   375
      Left            =   7560
      TabIndex        =   46
      ToolTipText     =   "Change size on draw field"
      Top             =   3960
      Width           =   2775
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   4695
      LargeChange     =   100
      Left            =   6960
      TabIndex        =   45
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll7 
      Height          =   255
      LargeChange     =   100
      Left            =   120
      TabIndex        =   44
      Top             =   4800
      Width           =   6855
   End
   Begin VB.PictureBox PC 
      Height          =   4695
      Left            =   120
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   453
      TabIndex        =   42
      Top             =   120
      Width           =   6855
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   7815
         Left            =   -360
         ScaleHeight     =   519
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   623
         TabIndex        =   43
         Top             =   -840
         Width           =   9375
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   375
      Left            =   3240
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   501
      TabIndex        =   27
      Top             =   6960
      Width           =   7575
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -15
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1
         TabIndex        =   29
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Output Colors"
      Height          =   1695
      Left            =   5400
      TabIndex        =   19
      Top             =   5160
      Width           =   2895
      Begin VB.CheckBox Check3 
         Caption         =   "Blue"
         Height          =   195
         Left            =   480
         TabIndex        =   34
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Green"
         Height          =   195
         Left            =   480
         TabIndex        =   33
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Red"
         Height          =   195
         Left            =   480
         TabIndex        =   32
         Top             =   480
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   30
         Text            =   "5"
         Top             =   720
         Width           =   375
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Use image colors"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         ScaleHeight     =   195
         ScaleWidth      =   1275
         TabIndex        =   22
         ToolTipText     =   "Click to change color"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Custom"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         Caption         =   "RGB (shaded)"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   1920
         X2              =   1320
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height sensivity"
         Height          =   255
         Left            =   1440
         TabIndex        =   31
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Position"
      Height          =   1695
      Left            =   2160
      TabIndex        =   11
      Top             =   5160
      Width           =   3135
      Begin VB.HScrollBar HScroll6 
         Height          =   255
         LargeChange     =   43
         Left            =   720
         Max             =   200
         TabIndex        =   39
         Top             =   1080
         Value           =   10
         Width           =   1815
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   255
         LargeChange     =   13
         Left            =   720
         Max             =   -30
         Min             =   30
         TabIndex        =   36
         Top             =   840
         Value           =   10
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Set default"
         Height          =   255
         Left            =   0
         TabIndex        =   18
         ToolTipText     =   "Set position values to default"
         Top             =   1440
         Width           =   3135
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         LargeChange     =   200
         Left            =   720
         Max             =   500
         Min             =   -500
         TabIndex        =   13
         Top             =   600
         Width           =   1815
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         LargeChange     =   200
         Left            =   720
         Max             =   500
         Min             =   -500
         TabIndex        =   12
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         Height          =   255
         Left            =   2520
         TabIndex        =   41
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Zoom"
         Height          =   255
         Left            =   0
         TabIndex        =   40
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         Height          =   255
         Left            =   2520
         TabIndex        =   38
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rotation"
         Height          =   255
         Left            =   0
         TabIndex        =   37
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Y"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "X"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input height"
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   1815
      Begin VB.OptionButton Option10 
         Caption         =   "Luminescence"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Invert"
         Height          =   195
         Left            =   1080
         TabIndex        =   50
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Saturation"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Hue"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   840
         Width           =   615
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   255
         LargeChange     =   40
         Left            =   120
         Max             =   1
         Min             =   100
         TabIndex        =   24
         Top             =   1920
         Value           =   5
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "All"
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   480
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Blue"
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Green"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Red"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "5"
         Height          =   255
         Left            =   1440
         TabIndex        =   26
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Height offset"
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   1680
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Draw heightmap!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      ToolTipText     =   "Draw heightmap with current settings"
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Import picture..."
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.VScrollBar VScroll1 
      Enabled         =   0   'False
      Height          =   2655
      LargeChange     =   50
      Left            =   10320
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   50
      Left            =   7320
      TabIndex        =   2
      Top             =   3240
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      Height          =   2655
      Left            =   7320
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   0
      Top             =   600
      Width           =   3015
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   0
         ScaleHeight     =   177
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   201
         TabIndex        =   1
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   2040
      TabIndex        =   28
      Top             =   6960
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PX As Integer
Dim PY As Integer
Dim Incrasement
Dim lefta
Dim XX As Integer
Dim YY As Integer

Dim CurX
Dim CurY

Dim JB As Byte

Dim Pixel

Dim lngBlue As Long
Dim lngGreen As Long
Dim lngRed As Long
Dim TempGrn As Long

Dim Hihi As Integer

Dim stoploop As Byte

Dim SB

Dim RRR As Integer
Dim GGG As Integer
Dim BBB As Integer

Dim Zoom

'0 = RGB, 1 = HSL
Dim RGBHSL As Byte
'0 = Dots, 1 = Connected lines
Dim DS As Byte

Dim HH As Integer
Dim SS As Integer
Dim LL As Integer

Dim Xinc As Integer
Dim Yinc As Integer

Dim SWD As Byte
Sub FadeColor()
On Error Resume Next

RRR = 0
GGG = 0
BBB = 0

If Check1.Value = 1 Then RRR = Hihi * Text1.Text
If Check2.Value = 1 Then GGG = Hihi * Text1.Text
If Check3.Value = 1 Then BBB = Hihi * Text1.Text

Pixel = RGB(RRR, GGG, BBB)


End Sub


Sub GetRGB(colors As String)


        lngBlue = Format(colors \ (16 ^ 4), "#00")
        
        'Get remainder value after Blue is taken out.
        TempGrn = colors Mod (16 ^ 4)
        
        'Use integer division to drop any decimal value
        lngGreen = Format(TempGrn \ (16 ^ 2), "#00")
        
        'Remainder is Red value
        lngRed = Format(TempGrn Mod (16 ^ 2), "#00")
        
        
        
End Sub

Sub GetHSL(colors As String)
   RGBtoHSL (colors)
End Sub


Sub SetH(AHH As Integer)
HH = AHH
End Sub
Sub SetHeightOffset()
On Error GoTo haha
'use red values
If Option1.Value = True Then Hihi = lngRed / HScroll4.Value
'use green values
If Option2.Value = True Then Hihi = lngGreen / HScroll4.Value
'use blue values
If Option3.Value = True Then Hihi = lngBlue / HScroll4.Value
'use all values
If Option4.Value = True Then
Hihi = Val(lngGreen + lngRed + lngBlue) / 3
Hihi = Hihi / HScroll4.Value
End If


If Check4.Value = 1 Then
Hihi = Hihi - Hihi - Hihi
End If

Exit Sub
haha:
MsgBox "Height offset value too high!", vbCritical
stoploop = 1
Exit Sub
End Sub

Sub SetHeightOffsetHSL()
On Error GoTo haha

'use red values
If Option6.Value = True Then Hihi = HH / HScroll4.Value
'use green values
If Option7.Value = True Then Hihi = SS / HScroll4.Value
'use blue values
If Option10.Value = True Then Hihi = LL / HScroll4.Value
'use all values



If Check4.Value = 1 Then
Hihi = Hihi - Hihi - Hihi
End If

Exit Sub
haha:
MsgBox "Height offset value too high!", vbCritical
stoploop = 1
Exit Sub
End Sub
Sub SetL(ALL As Integer)
LL = ALL
End Sub

Sub SetS(ASS As Integer)
SS = ASS
End Sub
Sub UppdateScroll()
'On Error Resume Next

If Picture3.Width <= Picture2.ScaleWidth Then
HScroll1.Enabled = False
Else
HScroll1.Enabled = True
End If

If Picture3.Height <= Picture2.ScaleHeight Then
VScroll1.Enabled = False
Else
VScroll1.Enabled = True
End If

VScroll1.Max = Picture3.ScaleHeight - Picture2.ScaleHeight
HScroll1.Max = Picture3.ScaleWidth - Picture2.ScaleWidth


End Sub
Sub GetMapInfo()
On Error GoTo jajaja
Picture1.Cls
If Form2.Check1.Value = 1 Then t1 = Timer
Label7.Caption = "Drawing..."
SB = Picture5.ScaleWidth / Picture3.ScaleHeight

Label7.Refresh
'Incrasement = Picture3.ScaleHeight * 0.01


SWD = Form2.Check2.Value

Xinc = Form2.Text1.Text
Yinc = Form2.Text2.Text

If Option1.Value = True Or Option2.Value = True Or Option3.Value = True Or Option4.Value = True Then RGBHSL = 0
If Option6.Value = True Or Option7.Value = True Or Option10.Value = True Then RGBHSL = 1

If Option11.Value = True Then DS = 0
If Option12.Value = True Then DS = 1
If Option13.Value = True Then DS = 2

If Option12.Value = True Or Option13.Value = True Then
Picture1.DrawWidth = Text2.Text
Else
Picture1.DrawWidth = 1
End If


XX = 0
YY = 0

Zoom = HScroll6.Value / 10
lefta = 0

Incrasement = 1

WWW = Val(Picture1.ScaleWidth / 2 + HScroll2.Value)
HHH = Val(Picture1.ScaleHeight / 2 - Picture3.ScaleHeight / 2 + HScroll3.Value)

LX = 0
LY = 0
Do

If stoploop = 1 Then
stoploop = 0
Exit Sub
End If

Pixel = GetPixel(Picture3.hdc, XX, YY)


If RGBHSL = 0 Then GetRGB (Pixel)
If RGBHSL = 1 Then GetHSL (Pixel)

If RGBHSL = 0 Then SetHeightOffset
If RGBHSL = 1 Then SetHeightOffsetHSL

If Option8.Value = True Then
Pixel = Picture4.BackColor
End If

If Option5.Value = True Then
FadeColor
End If

'Dot
If DS = 0 Then SetPixelV Picture1.hdc, Val(WWW + XX * Zoom - lefta), Val(HHH + YY * Zoom) - Hihi * Zoom, Pixel
'Spikes
If DS = 2 Then
Picture1.ForeColor = Pixel
Picture1.Line (Val(WWW + XX * Zoom - lefta), Val(HHH + YY * Zoom))-(Val(WWW + XX * Zoom - lefta), Val(HHH + YY * Zoom) - Hihi * Zoom)
End If
'Connected line
If DS = 1 Then
Picture1.ForeColor = Pixel
If LX = 0 Then
LX = Val(WWW + XX * Zoom - lefta)
LY = Val(HHH + YY * Zoom) - Hihi * Zoom
End If
Picture1.Line (LX, LY)-(Val(WWW + XX * Zoom - lefta), Val(HHH + YY * Zoom) - Hihi * Zoom)
LX = Val(WWW + XX * Zoom - lefta)
LY = Val(HHH + YY * Zoom) - Hihi * Zoom
End If


XX = XX + Xinc

If XX >= Picture3.ScaleWidth Then
XX = 0
YY = YY + Yinc
lefta = lefta + HScroll5.Value / 10 * Zoom
Picture6.Width = Picture6.Width + SB
Picture6.Refresh
If SWD = 1 Then Picture1.Refresh

If DS = 1 Then
LX = 0
LY = 0
End If
End If

Loop Until YY >= Picture3.ScaleHeight

Picture6.Width = 1
If Form2.Check1.Value = 1 Then
Label7.Caption = "Completed in " & Format(Timer - t1, "##.000") & " sec"
Else
Label7.Caption = ""
End If

Picture1.Refresh
Exit Sub
jajaja:
Picture6.Width = 1
Picture1.Refresh
Label7.Caption = "Error!"
MsgBox "Error while drawing heightmap, please check values!", vbCritical
Exit Sub
End Sub


Sub UppdateScroll2()
'Heightmap


If Picture1.Width <= PC.ScaleWidth Then
HScroll7.Enabled = False
Else
HScroll7.Enabled = True
End If

If Picture1.Height <= PC.ScaleHeight Then
VScroll2.Enabled = False
Else
VScroll2.Enabled = True
End If

VScroll2.Max = Picture1.ScaleHeight - PC.ScaleHeight
HScroll7.Max = Picture1.ScaleWidth - PC.ScaleWidth

Command6.Value = True
End Sub

Private Sub Command1_Click()
PopupMenu MenuForm.Menuf
End Sub

Private Sub Command2_Click()
stoploop = 0
GetMapInfo
End Sub


Private Sub Command3_Click()
HScroll2.Value = 0
HScroll3.Value = 0
HScroll5.Value = 10
HScroll6.Value = 10
End Sub

Private Sub Command4_Click()
On Error GoTo jaja
CM.CancelError = True
CM.Filter = "Bitmap (bmp)|*.bmp"
CM.ShowSave
SavePicture Picture1.Image, CM.FileName
MsgBox "Picture saved!", vbInformation
Exit Sub
jaja:
Exit Sub
End Sub

Private Sub Command5_Click()
On Error GoTo kaka
AX = InputBox("Enter new field width (max 2000)", "Field size", Picture1.ScaleWidth)
If AX = "" Then Exit Sub
If AX <= 10 Or AX > 2000 Then
MsgBox "Value to small/big", vbCritical
Exit Sub
End If

AY = InputBox("Enter new field height (max 2000)", "Field size", Picture1.ScaleHeight)
If AY = "" Then Exit Sub
If AY <= 10 Or AY > 2000 Then
MsgBox "Value to small/big", vbCritical
Exit Sub
End If

Picture1.Width = AX
Picture1.Height = AY

UppdateScroll2

Exit Sub
kaka:
MsgBox "Bad value!", vbCritical
Exit Sub
End Sub

Private Sub Command6_Click()
On Error Resume Next
If HScroll7.Enabled = True Then HScroll7.Value = HScroll7.Max / 2
If VScroll2.Enabled = True Then VScroll2.Value = VScroll2.Max / 2
End Sub


Private Sub Command7_Click()
On Error Resume Next
If HScroll1.Enabled = True Then HScroll1.Value = 0
If VScroll1.Enabled = True Then VScroll1.Value = 0
End Sub

Private Sub Command8_Click()
Form2.Show
End Sub

Private Sub Command9_Click()
MsgBox "Heightmap. Copyright (C) Johannes B 2002. Thanks to Alfred Koppold for pcx and tga file reading", vbInformation
End Sub

Private Sub Form_Load()

Form2.Show
Form2.Hide

UppdateScroll

Picture1.Width = 800
Picture1.Height = 800

UppdateScroll2
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub HScroll1_Change()
Picture3.Left = 0 - HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
HScroll1_Change
End Sub


Private Sub HScroll2_Change()
Label1.Caption = HScroll2.Value
End Sub

Private Sub HScroll3_Change()
Label2.Caption = HScroll3.Value
End Sub


Private Sub HScroll4_Change()
Label6.Caption = HScroll4.Value
End Sub

Private Sub HScroll5_Change()
Label10.Caption = HScroll5.Value
End Sub

Private Sub HScroll6_Change()
Label12.Caption = HScroll6.Value
End Sub

Private Sub HScroll7_Change()
Picture1.Left = 0 - HScroll7.Value
End Sub

Private Sub HScroll7_Scroll()
HScroll7_Change
End Sub


Private Sub Option11_Click()
Text2.Enabled = False
End Sub

Private Sub Option12_Click()
Text2.Enabled = True
End Sub


Private Sub Option13_Click()
Text2.Enabled = True
End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
JB = 1
CurX = X
CurY = Y
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If JB = 1 Then
If HScroll7.Enabled = True Then HScroll7.Value = HScroll7.Value + CurX - X
If VScroll2.Enabled = True Then VScroll2.Value = VScroll2.Value + CurY - Y
End If
End Sub


Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
JB = 0
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
JB = 1
CurX = X
CurY = Y

End Sub


Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If JB = 1 Then
If HScroll1.Enabled = True Then HScroll1.Value = HScroll1.Value + CurX - X
If VScroll1.Enabled = True Then VScroll1.Value = VScroll1.Value + CurY - Y
End If
End Sub


Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
JB = 0

End Sub

Private Sub Picture4_Click()
On Error GoTo haha
CM.CancelError = True
CM.ShowColor
Picture4.BackColor = CM.Color
Exit Sub
haha:
Exit Sub
End Sub

Private Sub VScroll1_Change()
Picture3.Top = 0 - VScroll1.Value
End Sub


Private Sub VScroll1_Scroll()
VScroll1_Change
End Sub


Private Sub VScroll2_Change()
Picture1.Top = 0 - VScroll2.Value
End Sub


Private Sub VScroll2_Scroll()
VScroll2_Change
End Sub


