VERSION 5.00
Begin VB.Form gen_parametros1 
   Caption         =   "Definir Cuentas"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12690
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   12690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Buscador de Cuentas"
      Height          =   375
      Left            =   360
      TabIndex        =   31
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Frame Frame9 
      Caption         =   "Cuentas Predefinidas STOCK"
      Height          =   1575
      Left            =   120
      TabIndex        =   28
      Top             =   3480
      Width           =   6255
      Begin VB.ComboBox c_resultado 
         Height          =   315
         Left            =   1920
         TabIndex        =   34
         Top             =   960
         Width           =   4215
      End
      Begin VB.ComboBox c_costo 
         Height          =   315
         Left            =   1920
         TabIndex        =   32
         Top             =   600
         Width           =   4215
      End
      Begin VB.ComboBox c_inventario 
         Height          =   315
         Left            =   1920
         TabIndex        =   29
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label27 
         BackColor       =   &H00C00000&
         Caption         =   "Resultados por tenencia"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C00000&
         Caption         =   "Costo Merc."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C00000&
         Caption         =   "Inventario  o Merc."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cuentas Contables Predefinidas COMPRAS"
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   6255
      Begin VB.ComboBox c_compvarias 
         Height          =   315
         Left            =   1920
         TabIndex        =   26
         Top             =   2040
         Width           =   4215
      End
      Begin VB.ComboBox c_nograb 
         Height          =   315
         Left            =   1920
         TabIndex        =   20
         Top             =   1680
         Width           =   4215
      End
      Begin VB.ComboBox c_retgan 
         Height          =   315
         Left            =   1920
         TabIndex        =   16
         Top             =   1320
         Width           =   4215
      End
      Begin VB.ComboBox c_retib 
         Height          =   315
         Left            =   1920
         TabIndex        =   14
         Top             =   960
         Width           =   4215
      End
      Begin VB.ComboBox c_ivac 
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   600
         Width           =   4215
      End
      Begin VB.ComboBox c_acreedores 
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C00000&
         Caption         =   "Compras Varias"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C00000&
         Caption         =   "Importes no Grabados"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         Caption         =   "Retenciones Gananc."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C00000&
         Caption         =   "Retenciones I.B.B.A"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         Caption         =   "Iva Compras"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Cuenta Acreedores"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   9960
      TabIndex        =   3
      Top             =   7680
      Width           =   1815
      Begin VB.CommandButton Command2 
         Caption         =   "Volver"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cuentas Contables Predefinidas VENTAS"
      Height          =   4815
      Left            =   6480
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin VB.ComboBox c_retsussv 
         Height          =   315
         Left            =   1800
         TabIndex        =   46
         Top             =   4320
         Width           =   4215
      End
      Begin VB.ComboBox c_retganv 
         Height          =   315
         Left            =   1800
         TabIndex        =   44
         Top             =   3960
         Width           =   4215
      End
      Begin VB.ComboBox c_retibbav 
         Height          =   315
         Left            =   1800
         TabIndex        =   41
         Top             =   3240
         Width           =   4215
      End
      Begin VB.ComboBox c_retivav 
         Height          =   315
         Left            =   1800
         TabIndex        =   40
         Top             =   3600
         Width           =   4215
      End
      Begin VB.ComboBox c_percsuss 
         Height          =   315
         Left            =   1800
         TabIndex        =   38
         Top             =   2880
         Width           =   4215
      End
      Begin VB.ComboBox c_percgan 
         Height          =   315
         Left            =   1800
         TabIndex        =   36
         Top             =   2520
         Width           =   4215
      End
      Begin VB.ComboBox c_perciva 
         Height          =   315
         Left            =   1800
         TabIndex        =   24
         Top             =   2160
         Width           =   4215
      End
      Begin VB.ComboBox C_percib 
         Height          =   315
         Left            =   1800
         TabIndex        =   22
         Top             =   1800
         Width           =   4215
      End
      Begin VB.ComboBox c_ivav 
         Height          =   315
         Left            =   1800
         TabIndex        =   18
         Top             =   1440
         Width           =   4215
      End
      Begin VB.ComboBox c_caja 
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Top             =   1080
         Width           =   4215
      End
      Begin VB.ComboBox c_ventas 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   720
         Width           =   4215
      End
      Begin VB.ComboBox c_deudores 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C00000&
         Caption         =   "Retencion Suss"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C00000&
         Caption         =   "Retencion Gan"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C00000&
         Caption         =   "Retencion I.B.B.A"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C00000&
         Caption         =   "Retencion Iva"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Percepcion Suss"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Percepcion Gan"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C00000&
         Caption         =   "Percepcion Iva"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C00000&
         Caption         =   "Percepcion I.B.B.A"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C00000&
         Caption         =   "Cuenta Iva Ventas"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C00000&
         Caption         =   "Cuenta Caja"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Cuenta Ventas"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Cuenta Deudores"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "gen_parametros1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private Sub c_acreedores_LostFocus()
If c_acreedores.ListIndex < 0 Then
     c_acreedores.ListIndex = 0
  End If
End Sub

Private Sub c_caja_LostFocus()
If c_caja.ListIndex < 0 Then
     c_caja.ListIndex = 0
  End If

End Sub

Private Sub c_costo_LostFocus()
If c_costo.ListIndex < 0 Then
     c_costo.ListIndex = 0
  End If
End Sub

Private Sub c_deudores_LostFocus()
If c_deudores.ListIndex < 0 Then
     c_deudores.ListIndex = 0
  End If

End Sub

Private Sub c_inventario_Change()
If c_inventario.ListIndex < 0 Then
     c_inventario.ListIndex = 0
End If
End Sub


Private Sub c_ventas_LostFocus()
If c_ventas.ListIndex < 0 Then
     c_ventas.ListIndex = 0
  End If

End Sub

Private Sub Command2_Click()
gen_parametros.Show
Me.Hide
End Sub

Private Sub Command3_Click()
cgr_buscacuenta.Show
End Sub

Private Sub Form_Load()
Call carga_cuentas_cont(c_deudores, "C", "D")
Call carga_cuentas_cont(c_acreedores, "C", "D")
Call carga_cuentas_cont(c_ventas, "C", "D")
Call carga_cuentas_cont(c_caja, "C", "D")

Call carga_cuentas_cont(c_ivac, "C", "D")
Call carga_cuentas_cont(c_retib, "C", "D")
Call carga_cuentas_cont(c_retgan, "C", "D")
Call carga_cuentas_cont(c_nograb, "C", "D")
Call carga_cuentas_cont(c_ivav, "C", "D")
Call carga_cuentas_cont(C_percib, "C", "D")
Call carga_cuentas_cont(c_perciva, "C", "D")
Call carga_cuentas_cont(c_compvarias, "C", "D")
Call carga_cuentas_cont(c_inventario, "C", "D")
Call carga_cuentas_cont(c_costo, "C", "D")
Call carga_cuentas_cont(c_resultado, "C", "D")

Call carga_cuentas_cont(c_retibbav, "C", "D")
Call carga_cuentas_cont(c_retivav, "C", "D")
Call carga_cuentas_cont(c_retganv, "C", "D")
Call carga_cuentas_cont(c_retsussv, "C", "D")



c_deudores.ListIndex = buscaindice(c_deudores, para.cuenta_deudores)
c_acreedores.ListIndex = buscaindice(c_acreedores, para.cuenta_acreedores)
c_ventas.ListIndex = buscaindice(c_ventas, para.cuenta_ventas)
c_caja.ListIndex = buscaindice(c_caja, para.cuenta_caja)

c_ivac.ListIndex = buscaindice(c_ivac, para.cuenta_iva_compras)
c_retib.ListIndex = buscaindice(c_retib, para.cuenta_retib)
c_retgan.ListIndex = buscaindice(c_retgan, para.cuenta_retgan)
c_nograb.ListIndex = buscaindice(c_nograb, para.cuenta_conceptos_nograbados)
c_ivav.ListIndex = buscaindice(c_ivav, para.cuenta_iva_ventas)
C_percib.ListIndex = buscaindice(C_percib, para.cuenta_perc_IB)
c_perciva.ListIndex = buscaindice(c_perciva, para.cuenta_perc_iva)
c_compvarias.ListIndex = buscaindice(c_compvarias, para.cuenta_compras_varias)
c_inventario.ListIndex = buscaindice(c_inventario, para.cuenta_inventario)
c_costo.ListIndex = buscaindice(c_costo, para.cuenta_costo)


c_retibbav.ListIndex = buscaindice(c_retibbav, para.cuenta_retibbav)
c_retivav.ListIndex = buscaindice(c_retivav, para.cuenta_retivav)
c_retganv.ListIndex = buscaindice(c_retganv, para.cuenta_retganv)
c_retsussv.ListIndex = buscaindice(c_retsussv, para.cuenta_retsussv)



End Sub

