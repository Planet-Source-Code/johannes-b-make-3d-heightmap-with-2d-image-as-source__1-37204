VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MenuForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menuform"
   ClientHeight    =   1890
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CM 
      Left            =   1080
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu Menuf 
      Caption         =   "mmm"
      Begin VB.Menu standard 
         Caption         =   "jpg,gif,bmp,dib,wmf,emf,ico,cur"
      End
      Begin VB.Menu pcx 
         Caption         =   "pcx"
      End
      Begin VB.Menu tga 
         Caption         =   "tga"
      End
   End
End
Attribute VB_Name = "MenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub pcx_Click()
Dim pcxFile As New LoadPCX


On Error GoTo kakaa
CM.CancelError = True
CM.Filter = "Pcx image (pcx)|*.pcx"
CM.ShowOpen


Screen.MousePointer = 11

pcxFile.Autoscale = True

pcxFile.LoadPCX CM.FileName
If pcxFile.IsPCX = True Then 'is it a PCX-File?
'Form2.ScaleMode = 3
pcxFile.ScaleMode = 3

Form1.Picture3.Width = pcxFile.PCXWidth '/ Scaling
Form1.Picture3.Height = pcxFile.PCXHeight '/ Scaling
pcxFile.DrawPCX Form1.Picture3

End If

'Form1.Picture3.Width = Form1.Picture3.Width / 15
'Form1.Picture3.Height = Form1.Picture3.Height / 15

Call Form1.UppdateScroll
Form1.Command7.Value = True

Screen.MousePointer = 0
Exit Sub
kakaa:
Screen.MousePointer = 0
Exit Sub
End Sub

Private Sub standard_Click()
On Error GoTo kaka
CM.CancelError = True
CM.Filter = "Picture (jpg,gif,bmp,dib,wmf,emf,ico,cur)|*.bmp;*.dib;*.jpg;*.gif;*.ico;*.cur;*.wmf;*.emf"
CM.ShowOpen
Form1.Picture3.Picture = LoadPicture(CM.FileName)

Call Form1.UppdateScroll
Form1.Command7.Value = True
Exit Sub
kaka:
Exit Sub
End Sub

Private Sub tga_Click()
Dim tgaFile As New LoadTGA


On Error GoTo kakaaa
CM.CancelError = True
CM.Filter = "Tga image (tga)|*.tga"
CM.ShowOpen

Screen.MousePointer = 11

tgaFile.Autoscale = True


tgaFile.LoadTGA CM.FileName
If tgaFile.IsTGA = True Then 'is it a TGA-File?
Form1.Picture3.Width = tgaFile.TGAWidth '/ Scaling
Form1.Picture3.Height = tgaFile.TGAHeight '/ Scaling
tgaFile.DrawTGA Form1.Picture3
End If

'Form1.Picture3.Width = Form1.Picture3.Width / 15
'Form1.Picture3.Height = Form1.Picture3.Height / 15


Call Form1.UppdateScroll
Form1.Command7.Value = True
Screen.MousePointer = 0
Exit Sub
kakaaa:
Screen.MousePointer = 0
Exit Sub
End Sub


