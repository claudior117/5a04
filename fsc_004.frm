VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fsc_formapago 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Forma de pago"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Forma de Pago"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.TextBox t_total 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   345
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   7
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox t_ingresado 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   2
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox t_diferencia 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   345
         Left            =   9840
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   1
         Top             =   3360
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid msf2 
         Height          =   3015
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   5318
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "A Ingresar"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Ingresado"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   5
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         Caption         =   "Diferencia"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   8520
         TabIndex        =   4
         Top             =   3360
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4110
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   17639
            MinWidth        =   17639
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "fsc_formapago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
Call armagrid2
End Sub
Sub armagrid2()
msf2.clear
msf2.Rows = 1
msf2.Cols = 12
msf2.ColWidth(0) = 600
msf2.ColWidth(1) = 1200
msf2.ColWidth(2) = 1200
msf2.ColWidth(3) = 2500
msf2.ColWidth(4) = 1700
msf2.ColWidth(5) = 1700
msf2.ColWidth(6) = 1000
msf2.ColWidth(7) = 1000
msf2.ColWidth(8) = 1000
msf2.ColWidth(9) = 1000
msf2.ColWidth(10) = 1000
msf2.ColWidth(11) = 500

msf2.TextMatrix(0, 0) = "Cod."
msf2.TextMatrix(0, 1) = "Forma Pago"
msf2.TextMatrix(0, 2) = "Num.Cheque"
msf2.TextMatrix(0, 3) = "Detalle/Banco"
msf2.TextMatrix(0, 4) = "Sucursal"
msf2.TextMatrix(0, 5) = "Titular"
msf2.TextMatrix(0, 6) = "Importe"
msf2.TextMatrix(0, 7) = "Fecha Dif."
msf2.TextMatrix(0, 8) = "Num.Int."
msf2.TextMatrix(0, 9) = "Cuenta"
msf2.TextMatrix(0, 10) = "Operacion"
msf2.TextMatrix(0, 11) = "Cod"

t_diferencia = ""
t_ingresado = ""

End Sub

Private Sub msf2_GotFocus()
Me.StatusBar1.Panels.item(1) = "[F1] Ch.Terc.  - [F2] TRansferncias - [F3] Otras formas pago - [ENTER] Continua "
If msf2.Rows > 0 Then
  msf2.FocusRect = flexFocusNone
Else
  msf2.FocusRect = flexFocusLight
End If
t_ingresado = suma_msflexgrid(msf2, 6)
t_diferencia = Format$(Val(t_total) - Val(t_ingresado), "######0.00")

End Sub

Private Sub msf2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
  vta_recibo3.Show
  vta_recibo3.t_modulo = "Q"
End If

If KeyCode = vbKeyF2 Then
  vta_recibo4.Show
  vta_recibo4.t_modulo = "Q"
End If


If KeyCode = vbKeyF1 Then
  vta_recibo2.Show
  vta_recibo2.t_modulo = "Q"
End If


If KeyCode = vbKeyF9 Then
 ' If Val(total) <> Val(t_ingresado) Then
 '    MsgBox ("El total ingresado no coincide con el total del Comprobante")
 ' Else
 '    total.SetFocus
 ' End If
  
End If


If KeyCode = vbKeyF5 Then
 If msf2.Rows > 2 Then
    msf2.RemoveItem (msf2.Row)
 Else
   Call armagrid2
 End If
End If


End Sub

Private Sub msf2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   fsc_tique2.t_ingreso.SetFocus
   Me.Hide
End If
End Sub

Private Sub msf2_LostFocus()
t_ingresado = suma_msflexgrid(msf2, 6)
msf2.FocusRect = flexFocusLight
End Sub
