VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_clientes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATOS DEL CLIENTE"
   ClientHeight    =   7380
   ClientLeft      =   3420
   ClientTop       =   945
   ClientWidth     =   6210
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7380
   ScaleWidth      =   6210
   Begin VB.TextBox t_codfiscal2 
      Height          =   285
      Left            =   2160
      TabIndex        =   44
      Top             =   6360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox t_percibeib 
      Height          =   285
      Left            =   1320
      TabIndex        =   40
      Top             =   6720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox t_codfiscal 
      Height          =   285
      Left            =   1320
      TabIndex        =   36
      Top             =   6360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox t_idvend 
      Height          =   285
      Left            =   360
      TabIndex        =   35
      Top             =   6720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox t_letrafact 
      Height          =   285
      Left            =   360
      TabIndex        =   34
      Top             =   6360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   5895
      Left            =   240
      TabIndex        =   15
      Top             =   0
      Width           =   5895
      Begin VB.TextBox t_idproveedor 
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Saldos"
         Height          =   255
         Left            =   4560
         TabIndex        =   41
         Top             =   4680
         Width           =   1095
      End
      Begin VB.TextBox t_dirlocal 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   4200
         Width           =   4335
      End
      Begin VB.TextBox t_limite 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   5520
         Width           =   1695
      End
      Begin VB.ComboBox c_iva 
         Height          =   315
         Left            =   3120
         TabIndex        =   5
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox t_cotiz 
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   5160
         Width           =   1095
      End
      Begin VB.TextBox t_saldo2 
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   4680
         Width           =   1095
      End
      Begin VB.TextBox t_saldo1 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   4680
         Width           =   1095
      End
      Begin VB.TextBox t_email 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   3840
         Width           =   4335
      End
      Begin VB.TextBox t_te 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3360
         Width           =   3255
      End
      Begin VB.TextBox t_cp 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox t_provincia 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox t_iva 
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox t_id 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   20
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
      Begin VB.Label Label17 
         BackColor       =   &H00FF0000&
         Caption         =   "Id.Proveedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   43
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FF0000&
         Caption         =   "Dir. Local"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   39
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackColor       =   &H000000FF&
         Caption         =   "                Limite Credito:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "$ por Dolar"
         Height          =   255
         Left            =   4080
         TabIndex        =   33
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000FF&
         Caption         =   "Cotizacion Ajuste entre cuentas:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   5160
         Width           =   2655
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "U$s"
         Height          =   255
         Left            =   4080
         TabIndex        =   30
         Top             =   4680
         Width           =   375
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "$"
         Height          =   255
         Left            =   2400
         TabIndex        =   29
         Top             =   4680
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H000000FF&
         Caption         =   "Saldos:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF0000&
         Caption         =   "Email:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF0000&
         Caption         =   "Telefonos:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF0000&
         Caption         =   "CP:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF0000&
         Caption         =   "Provincia:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "Cuit:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Localidad:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Dir.  Fiscal"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Razon Social"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   4440
      TabIndex        =   12
      Top             =   6000
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta020.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "vta020.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   13
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
      TabIndex        =   11
      Top             =   7125
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
Attribute VB_Name = "vta_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim habilitafacturaremito As Boolean



Private Sub btnacepta_Click()
 Call cambiacli
 Me.Hide
End Sub

Private Sub btnsale_Click()
  Me.Hide
End Sub


Sub limpia()
     t_cli = "Contado"
     t_direccion = "Rojas"
     t_localidad = "Rojas"
     t_cuit = "0"
     t_cp = "2705"
     t_provincia = "B"
     t_email = " "
     t_te = " "
     t_saldo1 = "0.00"
     t_saldo2 = "0.00"
     c_iva.ListIndex = 2
     t_cotiz = "0.000"
     t_idvend = 1
     t_letrafact = "B"
     t_codfiscal = "F"
 

 End Sub

Private Sub c_iva_LostFocus()
Call cambiacli
End Sub
Sub cambiacli()
Set rs = New ADODB.Recordset
q = "select * from g3 where [cod_tipoiva] = " & c_iva.ItemData(c_iva.ListIndex)
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  t_letrafact = rs("letra_cliente")
  t_iva = rs("abreviatura")
  t_codfiscal = rs("cod_fiscal")
Else
  c_iva.ListIndex = 2
  t_iva = "C.F"
  t_letrafact = "B"
  t_codfiscal = "F"
End If
If c_iva.ItemData(c_iva.ListIndex) <> 3 Then
  'VERIFICO CUIT
  If verificacuit(t_cuit) = 0 Then
    MsgBox ("Error en el numero de Cuit")
    c_iva.ListIndex = 2
    t_iva = "C.F"
    t_letrafact = "B"
    t_codfiscal = "F"
   End If
End If
End Sub

Private Sub Command1_Click()
Set cl_cli = New Clientes
cl_cli.carga (Val(t_id))
If cl_cli.id > 1 Then
    t_saldo1 = Format$(cl_cli.saldo(True, Now, True), "######0.00")
     t_saldo2 = Format$(cl_cli.saldo(True, Now, False), "######0.00")
      If Val(t_saldo2) <> 0 Then
       t_cotiz = Format$(Val(t_saldo1) / Val(t_saldo2), "####0.000")
     Else
       t_cotiz = "1.000"
      End If
     Else
     t_saldo1 = "0.00"
     t_saldo2 = "0.00"
     t_cotiz = para.cotizacion
End If
Set cl_cli = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF9 Then
   Call cambiacli
   Me.Hide
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call TabEnter2(Me, 10)
End If

End Sub

Sub carga()
If Val(t_id) > 0 Then
   Set cl_cli = New Clientes
   cl_cli.carga (Val(t_id))
   If cl_cli.id > 0 Then
     t_cli = cl_cli.razonsocial
     t_direccion = cl_cli.direccion
     t_localidad = cl_cli.localidad
     t_cuit = cl_cli.CUIT
     t_cp = cl_cli.cp
     t_provincia = cl_cli.provincia
     t_email = cl_cli.email
     t_te = cl_cli.te
     t_iva = cl_cli.abreviatura_tipoiva
     c_iva.ListIndex = buscaindice(c_iva, cl_cli.idtipoiva)
     t_idvend = cl_cli.idvendedor
     t_letrafact = cl_cli.letra
     t_codfiscal = cl_cli.codfiscal
     t_codfiscal2 = cl_cli.codfiscal2
     t_dirlocal = cl_cli.direccion_local
     t_percibeib = cl_cli.perciveib
     t_idproveedor = cl_cli.idproveedor
    ' If cl_cli.id > 1 Then
      't_saldo1 = Format$(cl_cli.saldo(True, Now, True), "######0.00")
     ' t_saldo2 = Format$(cl_cli.saldo(True, Now, False), "######0.00")
     ' If Val(t_saldo2) <> 0 Then
     '  t_cotiz = Format$(Val(t_saldo1) / Val(t_saldo2), "####0.000")
    ' Else
    '   t_cotiz = "1.000"
    '  End If
    ' Else
    '  t_saldo1 = "0.00"
     ' t_saldo2 = "0.00"
    '  t_cotiz = para.cotizacion
   '  End If
     t_limite = cl_cli.limitecredito
   
   End If
   Set cl_cli = Nothing
End If
End Sub



Private Sub Form_Load()
Call carga_tipoiva(c_iva)
c_iva.ListIndex = 0
Me.StatusBar1.Panels.item(1) = "[F9] Cambia y sale - [ESC] Sale sin cambios - "

End Sub

Private Sub t_cuit_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)
End Sub

Private Sub t_cuit_LostFocus()
If t_cuit = "" Then
  t_cuit = "0"
End If
End Sub

Private Sub t_dirlocal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If
End Sub

