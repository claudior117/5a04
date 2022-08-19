VERSION 5.00
Begin VB.Form vta_listaprecios5 
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
Attribute VB_Name = "vta_listaprecios5"
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
   Select Case para.tipolistaprecios
    Case Is = 2
      vta_listaprecios_2.t_detalle = t_texto
      vta_listaprecios_2.carga
    Case Is = 3
      vta_listaprecios_3.t_detalle = t_texto
      vta_listaprecios_3.carga
     Case Is = 4
      vta_listaprecios_4.t_detalle = t_texto
      vta_listaprecios_4.carga
      Case Is = 5
      vta_listaprecios_5.t_detalle = t_texto
      vta_listaprecios_5.carga
    
    Case Else
      vta_listaprecios.t_detalle = t_texto
      vta_listaprecios.carga
   End Select
  End If
  Unload Me
End If


If KeyAscii = 27 Then
 Unload Me
End If

End Sub
