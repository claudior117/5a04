VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form com_proveedor 
   BackColor       =   &H00E0E0E0&
   Caption         =   "DATOS DEL PROVEEDOR"
   ClientHeight    =   6990
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   6210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6990
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_letra 
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   5415
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   5895
      Begin VB.ComboBox c_iva 
         Height          =   315
         Left            =   3120
         TabIndex        =   5
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox t_cotiz 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox t_saldo2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   4440
         Width           =   1695
      End
      Begin VB.TextBox t_saldo1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox t_email 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   3960
         Width           =   4335
      End
      Begin VB.TextBox t_te 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   3480
         Width           =   3255
      End
      Begin VB.TextBox t_cp 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox t_provincia 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox t_iva 
         Height          =   285
         Left            =   5160
         TabIndex        =   6
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox t_id 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox t_cuit 
         Height          =   285
         Left            =   1320
         MaxLength       =   11
         TabIndex        =   4
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox t_localidad 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   1560
         Width           =   4095
      End
      Begin VB.TextBox t_direccion 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox t_cli 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FF0000&
         Caption         =   "Cotizacion Ajuste entre cuentas:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   4920
         Width           =   2655
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "$ por Dolar"
         Height          =   255
         Left            =   4080
         TabIndex        =   31
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "U$s"
         Height          =   255
         Left            =   5280
         TabIndex        =   29
         Top             =   4440
         Width           =   375
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "$"
         Height          =   255
         Left            =   2880
         TabIndex        =   28
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF0000&
         Caption         =   "Saldos:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H000000FF&
         Caption         =   "Email:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H000000FF&
         Caption         =   "Telefonos:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H000000FF&
         Caption         =   "CP:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H000000FF&
         Caption         =   "Provincia:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000C0&
         Caption         =   "Id.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H000000FF&
         Caption         =   "Cuit:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000FF&
         Caption         =   "Localidad:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   "Direccion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   "Razon Social"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   4440
      TabIndex        =   11
      Top             =   5640
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "com001.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "com001.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   6735
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "com_proveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim habilitafacturaremito As Boolean



Private Sub btnacepta_Click()
Call cambiaprov
Me.Hide
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub


Sub limpia()
 t_cli = " "
 t_direccion = " "
 t_localidad = " "
 t_cuit = " "
 t_iva = " "
 t_provincia = " "
 t_cp = " "
 t_te = " "
 t_email = " "
 t_saldo1 = " "
 t_saldo2 = " "

 End Sub

Private Sub c_iva_LostFocus()
If c_iva.ListIndex < 0 Then
  c_iva.ListIndex = 0
End If
Call cambiaprov
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF9 Then

End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call TabEnter2(Me, 9)
End If

End Sub

Sub cambiaprov()
Set rs = New ADODB.Recordset
q = "select * from g3 where [cod_tipoiva] = " & c_iva.ItemData(c_iva.ListIndex)
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  t_letra = rs("letra_prov")
  t_iva = rs("abreviatura")
Else
  c_iva.ListIndex = 2
  t_iva = "C.F"
  t_letra = "C"
End If
End Sub

Sub carga()
Call limpia
If Val(t_id) > 0 Then
   Set cl_prov2 = New proveedores
   cl_prov2.carga (Val(t_id))
   If cl_prov2.idprov > 0 Then
     t_cli = cl_prov2.razonsocial
     t_direccion = cl_prov2.direccion
     t_localidad = cl_prov2.localidad
     t_cuit = cl_prov2.CUIT
     t_cp = cl_prov2.cp
     t_provincia = cl_prov2.provincia
     t_email = cl_prov2.email
     t_te = cl_prov2.te
     c_iva.ListIndex = buscaindice(c_iva, cl_prov2.codtipoiva)
     t_saldo1 = Format$(cl_prov2.saldo(True, Now, True, 0), "######0.00")
     t_saldo2 = Format$(cl_prov2.saldo(True, Now, False, 0), "######0.00")
     't_iva = cl_prov.abreviatura_tipoiva
     If Val(t_saldo2) <> 0 Then
       t_cotiz = Format$(Val(t_saldo1) / Val(t_saldo2), "####0.000")
     Else
       t_cotiz = "0.000"
     End If
   End If
   Set cl_prov2 = Nothing
End If
Call cambiaprov
End Sub



Private Sub Form_Load()
Call carga_tipoiva(c_iva)
c_iva.ListIndex = 0

End Sub

Private Sub t_email_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 btnacepta.SetFocus
End If
End Sub
