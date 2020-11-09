VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_gencot 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GENERACION DE CODIGO DE OPERACION DE TRASNPORTE(Archivo txt para SIAP)"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   8520
      TabIndex        =   18
      Top             =   120
      Width           =   3255
      Begin VB.TextBox t_idcomp 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   12
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Id. Comp"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   7095
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   8295
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Consumidor Final"
         Height          =   375
         Left            =   6840
         TabIndex        =   29
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox t_dpto 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7080
         MaxLength       =   150
         TabIndex        =   8
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox t_numero 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2880
         Width           =   975
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1440
         Width           =   4335
      End
      Begin VB.TextBox t_codloc 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   26
         Top             =   2160
         Width           =   735
      End
      Begin VB.ComboBox c_loc 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2160
         Width           =   4335
      End
      Begin VB.TextBox t_cuit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   1
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox t_piso 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4800
         MaxLength       =   150
         TabIndex        =   7
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox t_cp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox t_direccion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   5
         Top             =   2520
         Width           =   4335
      End
      Begin VB.TextBox t_id 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox t_descripcion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   0
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Dpto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   6000
         TabIndex        =   28
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Piso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   3720
         TabIndex        =   27
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Cod. Postal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   25
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Numero Doc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   480
         TabIndex        =   24
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Numero:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   23
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   22
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Localiadad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   21
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "DESTINO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Razon Social"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Calle "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   2520
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   10
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "vta027.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "vta027.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   11
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
      TabIndex        =   9
      Top             =   8265
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "14/04/2009"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "03:55 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_gencot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String



Private Sub btnsale_Click()
Unload Me
End Sub



Private Sub c_loc_LostFocus()
If c_loc.ListIndex < 0 Then
   c_loc.ListIndex = 0
End If
t_codloc = c_loc.ItemData(c_loc.ListIndex)

End Sub

Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
  c_prov.ListIndex = 0
End If

x = verificaloc

End Sub

Private Sub c_vend_LostFocus()
If c_vend.ListIndex < 0 Then
  c_vend.ListIndex = 0
End If
End Sub

Private Sub Form_Activate()
Call carga
End Sub
Function verificaloc() As Boolean
   Set rs = New ADODB.Recordset
   q = "select * from cot_01 where [cp] = " & Val(t_cp)
   rs.Open q, cn1
   If Not rs.EOF And Not rs.BOF Then
     If rs("id_provincia") = c_prov.ItemData(c_prov.ListIndex) Then
         verificaloc = True
     Else
        MsgBox ("El Codigo Postal No pertenece a la Provincia")
        verificaloc = False
     End If
   Else
     MsgBox ("El Codigo Postal No Existe")
     verificaloc = False
   End If
   Set rs = Nothing
End Function
Sub carga()
Set cl_compvta = New comprobantes_venta
cl_compvta.cargar2 (Val(t_idcomp))
      
Set cl_cli = New Clientes
cl_cli.carga (cl_compvta.idcliente)
t_descripcion = cl_cli.razonsocial
c_prov.ListIndex = buscaindice(c_prov, cl_cli.id_provincia)
t_cp = cl_cli.cp
If verificaloc Then
   Call cargaloccot(Val(t_cp))
End If
t_direccion = cl_cli.Direccion
If cl_cli.idtipoiva <> 3 Then
   t_cuit = cl_cli.CUIT
   Check1 = False
Else
   t_cuit = " "
   Check1 = True
End If
      
      
Set cl_compvta = Nothing
Set cl_cli = Nothing
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 15)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call carga_provincias(c_prov)

End Sub

Private Sub t_cp_LostFocus()
Call NULOS(t_cp)
x = verificaloc
Call cargaloccot(Val(t_cp))


End Sub
Sub cargaloccot(cp)
q = "select * from cot_01 where [cp] = " & cp & " order by [localidad]"
Set rs = New ADODB.Recordset
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
  c_prov.ListIndex = buscaindice(c_prov, rs("id_provincia"))
  Call llena_combo(rs, "localidad", "cod_localidad", c_loc, True)
  c_loc.ListIndex = 0
  t_codloc = c_loc.ItemData(c_loc.ListIndex)
Else
  MsgBox ("Codigo Postal Inexistente")
  c_loc.clear
  c_loc.AddItem "<ERROR>"
  t_codloc = 0
End If
c_loc.ListIndex = 0
Set rs = Nothing

End Sub
Private Sub t_credito_LostFocus()
t_credito = Format$(Val(t_credito), "######0.00")
End Sub

Private Sub t_cuit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 btnacepta.SetFocus
End If
End Sub

Private Sub t_cuit_LostFocus()
Call NULOS(t_cuit)
End Sub

Private Sub t_descripcion_LostFocus()
Call NULOS(t_descripcion)
End Sub



Private Sub t_direccion_LostFocus()
Call NULOS(t_direccion)
End Sub

Private Sub t_email_LostFocus()
Call NULOS(t_email)
End Sub



Private Sub t_percib_LostFocus()
T_PERCIB = Format$(T_PERCIB, ">@")
Select Case T_PERCIB
Case Is = "S", Is = "N"

Case Else
  T_PERCIB = "S"
End Select

End Sub


Private Sub t_saldoi_LostFocus()
t_saldoi = Format$(t_saldoi, ">@")
Select Case t_saldoi
Case Is = "S", Is = "N"

Case Else
  t_saldoi = "N"
End Select

End Sub

