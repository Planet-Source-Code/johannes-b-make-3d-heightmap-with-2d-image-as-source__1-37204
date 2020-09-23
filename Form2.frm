VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced settings"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2880
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   2880
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CM 
      Left            =   2280
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   1800
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   10
      ToolTipText     =   "Click here to change"
      Top             =   2040
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   8
      ToolTipText     =   "Click here to change"
      Top             =   1800
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Show while drawing (slower)"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show time taken after drawing"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "1"
      ToolTipText     =   "Incrase for faster drawing"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "1"
      ToolTipText     =   "Incrase for faster drawing"
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Progressbar color"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "HM Background color"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Warning. Changing some of these values to much may cause program to hang"
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Y incrasement"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "X incrasement"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HH As Byte
Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Load()
HH = 0
End Sub

Private Sub Picture1_Click()
On Error GoTo hahag
CM.CancelError = True
If HH = 0 Then MsgBox "Warning. If a current heightmap is drawed it will be cleared if you chose a new color. Click OK then Cancel to keep heightmap/backcolor", vbExclamation
HH = 1
CM.ShowColor
Picture1.BackColor = CM.Color
Form1.Picture1.BackColor = CM.Color
Exit Sub
hahag:
Exit Sub
End Sub

Private Sub Picture2_Click()
On Error GoTo hahae
CM.CancelError = True
CM.ShowColor
Picture2.BackColor = CM.Color
Form1.Picture6.BackColor = CM.Color
Exit Sub
hahae:
Exit Sub
End Sub


