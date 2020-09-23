VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmOil 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "OilPainting"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   478
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   473
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox PicFront 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   6780
      Left            =   0
      ScaleHeight     =   452
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   532
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   7980
   End
   Begin VB.HScrollBar ScrSmooth 
      Height          =   255
      Left            =   4680
      Max             =   255
      TabIndex        =   4
      Top             =   60
      Value           =   128
      Width           =   2355
   End
   Begin VB.HScrollBar ScrBr 
      Height          =   255
      Left            =   2520
      Max             =   5
      Min             =   1
      TabIndex        =   2
      Top             =   60
      Value           =   3
      Width           =   1275
   End
   Begin VB.CommandButton CmdLoad 
      Caption         =   "Load"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   3780
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   6615
      Left            =   0
      ScaleHeight     =   441
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   533
      TabIndex        =   0
      Top             =   420
      Width           =   7995
   End
   Begin VB.Label Label2 
      Caption         =   "Smooth"
      Height          =   255
      Left            =   4020
      TabIndex        =   5
      Top             =   60
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Brush Size"
      Height          =   255
      Left            =   1620
      TabIndex        =   3
      Top             =   60
      Width           =   795
   End
End
Attribute VB_Name = "FrmOil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'OilPaintig demo
'© Scythe 2003
'This a part of my FX Module created for LM-X

'Compile to see real speed

'More Infos in the BAS File


Private Sub Pic_Resize()
Dim x As Long
If Pic.Width < 470 Then
x = 470
Else
x = Pic.Width
End If

Me.Width = (x + (Me.Width / Screen.TwipsPerPixelX - Me.ScaleWidth)) * Screen.TwipsPerPixelX
Me.Height = (Pic.Top + Pic.Height + (Me.Height / Screen.TwipsPerPixelY - Me.ScaleHeight)) * Screen.TwipsPerPixelY

PicFront.Width = Pic.Width
PicFront.Height = Pic.Height

'Slow but it does what i need
PicFront.PaintPicture Pic.Image, 0, 0, Pic.Width, Pic.Height, 0, 0, Pic.Width, Pic.Height, vbSrcCopy
End Sub

Private Sub CmdLoad_Click()
On Error GoTo ErrOut

cmd.ShowOpen

If cmd.filename = "" Then Exit Sub
Pic.Picture = LoadPicture(cmd.filename)
ErrOut:
End Sub

Private Sub ScrBr_Change()
Me.MousePointer = 11
PicFront.PaintPicture Pic.Image, 0, 0, Pic.Width, Pic.Height, 0, 0, Pic.Width, Pic.Height, vbSrcCopy
PicOilPaint PicFront, 0, 0, PicFront.ScaleWidth - 1, PicFront.ScaleHeight - 1, ScrBr.Value, ScrSmooth.Value
Me.MousePointer = 0
PicFront.Refresh
End Sub


Private Sub ScrSmooth_Change()
Me.MousePointer = 11
PicFront.PaintPicture Pic.Image, 0, 0, Pic.Width, Pic.Height, 0, 0, Pic.Width, Pic.Height, vbSrcCopy
PicOilPaint PicFront, 0, 0, PicFront.ScaleWidth - 1, PicFront.ScaleHeight - 1, ScrBr.Value, ScrSmooth.Value
Me.MousePointer = 0
PicFront.Refresh
End Sub
