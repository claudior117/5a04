VERSION 5.00
Begin VB.Form gen_logo 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   735
      Left            =   3600
      ScaleHeight     =   675
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   1800
      Width           =   3375
   End
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   480
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   480
      ScaleHeight     =   1035
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "cod barra facetura e"
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "firma"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   2160
      Width           =   735
   End
End
Attribute VB_Name = "gen_logo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private Sub Form_Load()
imagen = App.Path & "\tools\logo.jpg"
Picture1.Picture = LoadPicture(App.Path & "\tools\logo.jpg")


End Sub

