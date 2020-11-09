VERSION 5.00
Begin VB.Form exp_productos1 
   Caption         =   "Busqueda de Productos"
   ClientHeight    =   750
   ClientLeft      =   4230
   ClientTop       =   4590
   ClientWidth     =   7290
   ControlBox      =   0   'False
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   750
   ScaleWidth      =   7290
   Begin VB.TextBox t_texto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "exp_productos1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub t_texto_GotFocus()
SendKeys ("{end}")

End Sub

Private Sub t_texto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If t_texto <> "" Then
    exp_productos.t_detalle = t_texto
    exp_productos.carga
  End If
  Unload Me
End If


If KeyAscii = 27 Then
 Unload Me
End If

End Sub
