VERSION 5.00
Begin VB.Form abm_comp_compra3 
   Caption         =   "Graba Comprobante"
   ClientHeight    =   1875
   ClientLeft      =   4065
   ClientTop       =   3450
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   5160
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de Actualizacion de estructura de Precios"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.OptionButton Option3 
         Caption         =   "No actualiza Nada"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   1080
         Width           =   3615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Solo actualiza Precio de COMPRA"
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   720
         Width           =   3615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Actualiza Precio de COMPRA y de VENTA"
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
   End
End
Attribute VB_Name = "abm_comp_compra3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Select Case para.tipoactupreciocompcompra
  Case Is = 1
    Option1 = True
    Option1.SetFocus
  Case Is = 2
    Option2 = True
    Option2.SetFocus
  Case Is = 3
    Option3 = True
    Option3.SetFocus
  Case Else
    Option3 = True
    Option3.SetFocus
End Select
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call ABM_COMP_COMPRA.mensaje
J = MsgBox("Graba " & ABM_COMP_COMPRA.Label18, 4)
If J = 6 Then
 If verificaperiodog(t_fecha) = "C" Then
   MsgBox ("Periodo Cerrado. Imposible grabar operacion")
 Else
  If Option1 = True Then
     Call ABM_COMP_COMPRA.graba(1)
  
  Else
    If Option2 = True Then
     Call ABM_COMP_COMPRA.graba(2)
    Else
     Call ABM_COMP_COMPRA.graba(3)
    End If
  End If
 End If
 Unload Me
End If

End If

If KeyAscii = 27 Then
  Unload Me
End If
End Sub

