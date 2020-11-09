VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form espere 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   1440
      Picture         =   "ESPERE.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   3120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Procesando......"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   3615
      Left            =   120
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "espere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private Sub Form_Load()
Label1 = "Procesando......."
End Sub
