VERSION 5.00
Begin VB.Form com_consultaapoc 
   Caption         =   "Consulta por CUIT en el padron de Facturas Apocrifas"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "CONSULTAR"
      Height          =   615
      Left            =   4080
      TabIndex        =   2
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "INGRESE NRO. CUIT(sin guines)"
      Height          =   1215
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         MaxLength       =   11
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
   End
End
Attribute VB_Name = "com_consultaapoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Len(Text1) <> 11 Then
  MsgBox ("Formato de Cuit NO valido")
Else
  If buscafacturaapocrifa(Val(Text1)) Then
     MsgBox ("¡ATENCION! El cuit ESTA registrado en el padron de facturas Apocrifas")
  Else
     MsgBox ("Proveedor Correcto")
  End If
End If
End Sub

Private Sub Text1_Click()
Text1 = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)
End Sub
