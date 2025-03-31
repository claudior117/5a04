VERSION 5.00
Begin VB.Form fsc_tiqueNF2 
   Caption         =   "CIERRE DE TIQUE"
   ClientHeight    =   3915
   ClientLeft      =   3180
   ClientTop       =   2475
   ClientWidth     =   6540
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   6540
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   6135
      Begin VB.TextBox t_vuelto 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "VUELTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin VB.CommandButton Command1 
         Caption         =   "Forma Pago"
         Height          =   495
         Left            =   5280
         TabIndex        =   8
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox t_ingreso 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox T_TOTAL 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "INGRESO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "fsc_tiqueNF2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
fsc_formapago.T_TOTAL = T_TOTAL
fsc_formapago.t_modulo = 2
fsc_formapago.Show
End Sub

Private Sub Command1_LostFocus()
t_ingreso = fsc_formapago.t_ingresado
End Sub

Private Sub Form_Activate()
Frame2.Visible = False
t_ingreso = ""

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   T_TOTAL = Format$(Val(T_TOTAL), "#####0.00")
   t_vuelto = Format$(Val(t_ingreso) - Val(T_TOTAL), "#####0.00")
   Frame2.Visible = True
End If

If KeyAscii = 27 Then
   Me.Hide
End If
End Sub

Private Sub t_ingreso_Change()
   T_TOTAL = Format$(Val(T_TOTAL), "#####0.00")
   t_vuelto = Format$(Val(t_ingreso) - Val(T_TOTAL), "#####0.00")
   Frame2.Visible = True
End Sub

Private Sub t_ingreso_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
   If estadocaja(fsc_tiqueNF.t_fecha) = "A" Then
    If Val(fsc_formapago.t_diferencia) = 0 And fsc_formapago.msf2.Rows > 1 Then
       Call iniciagraba
    Else
       If fsc_formapago.msf2.Rows <= 1 Then
              'pone forma de pago efectivo
              fsc_formapago.msf2.AddItem "001" & Chr(9) & 1 & Chr(9) & "-" & Chr(9) & "Efectivo $" & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & Format$(Val(T_TOTAL), "######0.00") & Chr(9) & Format$(fsc_tiqueNF.t_fecha, "DD/MM/YYYY") & Chr(9) & "" & Chr(9) & para.cuenta_caja & Chr(9) & "Pago Efectivo"
              Call iniciagraba
       Else
          MsgBox ("El pago ingresado no coincide con el total del comprobante")
       End If
    End If
   Else
     MsgBox ("Caja Cerrada. Imposible ingresar movimientos de contado en la fecha indicada")
   End If
   
 End If

End Sub

Sub iniciagraba()
J = MsgBox("Cierra Tique (S/N", 4)
   If J = 6 Then
     
      fsc_tiqueNF.cierratique2
      fsc_tiqueNF.Label2 = "Cerrando Tique"
      Me.Hide
   
   End If

End Sub

