VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_transporte 
   BackColor       =   &H00E0E0E0&
   Caption         =   "SELECCION DE TRANSPORTE"
   ClientHeight    =   6810
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   7005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6810
   ScaleWidth      =   7005
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos para el REMITO"
      Height          =   3975
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   5895
      Begin VB.TextBox t_dni 
         Height          =   285
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   6
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox t_acoplado 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   8
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox t_dominio 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   7
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox t_chofer 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   5
         Top             =   2280
         Width           =   4095
      End
      Begin VB.TextBox t_id 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox t_cuit 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox t_localidad 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   1560
         Width           =   4095
      End
      Begin VB.TextBox t_direccion 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox t_transp 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FF0000&
         Caption         =   "DNI :"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF0000&
         Caption         =   "Dominio Acop."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF0000&
         Caption         =   "Dominio :"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF0000&
         Caption         =   "Chofer"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000C0&
         Caption         =   "Id.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "Cuit:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Localidad:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Direccion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Transporte:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TRANSPORTES // CAMIONES registrados en el Sistema "
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   6735
      Begin VB.ComboBox c_camion 
         Height          =   315
         Left            =   1080
         TabIndex        =   27
         Text            =   "Combo1"
         Top             =   600
         Width           =   4695
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   5880
         Picture         =   "vta018.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   5880
         Picture         =   "vta018.frx":0105
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox C_TRANSP 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C000C0&
         Caption         =   "Camion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C000C0&
         Caption         =   "Transporte:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   4440
      TabIndex        =   11
      Top             =   5520
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta018.frx":020A
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
         Picture         =   "vta018.frx":0A8C
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
      Top             =   6555
      Width           =   7005
      _ExtentX        =   12356
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
Attribute VB_Name = "vta_transporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim habilitafacturaremito As Boolean



Private Sub btnacepta_Click()
Me.Hide
End Sub

Private Sub btnsale_Click()
Me.Hide
End Sub


Private Sub c_camion_LostFocus()
If c_camion.ListIndex <> 0 Then
  q = "select * from a17 where [id_camion] = " & c_camion.ItemData(c_camion.ListIndex)
  Set rs = New ADODB.Recordset
  rs.MaxRecords = 1
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
     'cargo datos
     T_chofer = rs("chofer")
     t_dni = rs("dni")
     t_dominio = rs("dominio")
     t_acoplado = rs("dominio_acoplado")
  End If
  Set rs = Nothing
End If

End Sub

Private Sub c_transp_LostFocus()
  Call buscadatos
End Sub
Sub buscadatos()
If c_transp.ListIndex <= 0 Then
   Call limpia
Else
   Set cl_prov = New proveedores
   cl_prov.carga (c_transp.ItemData(c_transp.ListIndex))
   If cl_prov.idprov > 0 Then
     Call carga_camiones(c_camion, cl_prov.idprov)
     c_camion.AddItem "<Sin detallar>", 0
     c_camion.ListIndex = 0
     t_transp = cl_prov.razonsocial
     t_direccion = cl_prov.direccion
     t_localidad = cl_prov.localidad
     t_cuit = cl_prov.CUIT
     t_id = cl_prov.idprov
   Else
     Call limpia
   End If
   Set cl_prov = Nothing
End If

End Sub
Sub limpia()
 t_transp = " "
 t_direccion = " "
 t_localidad = " "
 t_cuit = " "
 T_chofer = " "
 t_dominio = " "
 t_acoplado = " "
 t_id = 1
 c_transp.ListIndex = 0
 
End Sub

Private Sub Command1_Click()
ABM_PROv.Show
End Sub

Private Sub Command1_LostFocus()
Call carga_transporte(c_transp)
c_transp.AddItem "<Sin ingresar>", 0
c_transp.ListIndex = 0
End Sub

Private Sub Command2_Click()
GEN_ABMCAMION.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call TabEnter2(Me, 8)
End If

End Sub

Private Sub Form_Load()
Call carga_transporte(c_transp)
c_transp.AddItem "<Sin ingresar>", 0
c_transp.ListIndex = 0
c_camion.AddItem "<Sin detallar>", 0
c_camion.ListIndex = 0


End Sub




Private Sub t_acoplado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 btnacepta.SetFocus
End If
End Sub

