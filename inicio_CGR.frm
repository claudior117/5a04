VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form inicio_CGR 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MODULO  CONTABILIDAD"
   ClientHeight    =   8415
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   12495
   FontTransparent =   0   'False
   Icon            =   "inicio_CGR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8415
   ScaleWidth      =   12495
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impresora Actual del Sistema"
      Height          =   615
      Left            =   4920
      TabIndex        =   26
      Top             =   7200
      Width           =   4815
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label7"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modulo "
      Height          =   1455
      Left            =   6600
      TabIndex        =   23
      Top             =   5520
      Width           =   2055
      Begin VB.Image Image1 
         Height          =   480
         Left            =   720
         Picture         =   "inicio_CGR.frx":030A
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "CONTABILIDAD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informes"
      Height          =   3015
      Left            =   6240
      TabIndex        =   20
      Top             =   1200
      Width           =   2415
      Begin VB.CommandButton Command7 
         Height          =   615
         Left            =   360
         Picture         =   "inicio_CGR.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Height          =   615
         Left            =   360
         Picture         =   "inicio_CGR.frx":0EF3
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Height          =   615
         Left            =   360
         Picture         =   "inicio_CGR.frx":1830
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1200
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Periodo Contable Actual"
      Height          =   735
      Left            =   6480
      TabIndex        =   16
      Top             =   120
      Width           =   5295
      Begin VB.ComboBox c_periodo 
         Height          =   360
         Left            =   120
         TabIndex        =   18
         Text            =   "Combo1"
         Top             =   240
         Width           =   4095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cambiar"
         Height          =   375
         Left            =   4320
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Contabilidad"
      Height          =   3015
      Left            =   3240
      TabIndex        =   13
      Top             =   1200
      Width           =   2415
      Begin VB.CommandButton Command6 
         Height          =   615
         Left            =   360
         Picture         =   "inicio_CGR.frx":2193
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Height          =   615
         Left            =   360
         Picture         =   "inicio_CGR.frx":2AF1
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   360
         Picture         =   "inicio_CGR.frx":3491
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   240
      TabIndex        =   4
      Top             =   6120
      Width           =   4575
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800080&
         Caption         =   "CUIT:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800080&
         Caption         =   "Telefono:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         Caption         =   "Direccion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800080&
         Caption         =   "Razon Social:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9960
      TabIndex        =   1
      Top             =   6840
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "inicio_CGR.frx":3D9F
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "inicio_CGR.frx":4621
         Style           =   1  'Graphical
         TabIndex        =   2
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
      TabIndex        =   0
      Top             =   8160
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   13229
            MinWidth        =   13229
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "16/07/2016"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:45 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Menu M_tablas 
      Caption         =   "&Tablas"
      Begin VB.Menu M_plnac 
         Caption         =   "Plan de Cuentas"
      End
      Begin VB.Menu M_periodos 
         Caption         =   "Periodos Contables"
      End
   End
   Begin VB.Menu M_consultas 
      Caption         =   "Herramientas"
      Begin VB.Menu M_asientos 
         Caption         =   "Agrupar Asientos"
      End
   End
   Begin VB.Menu M_jhgd 
      Caption         =   "Asientos Automaticos Provisorios"
      Begin VB.Menu M_asa 
         Caption         =   "Asientos Automaticos Generados"
      End
      Begin VB.Menu M_maau 
         Caption         =   "Mayores de Asientos Automaticos"
      End
      Begin VB.Menu M_admasit 
         Caption         =   "Administrador de Asientos Temporales"
      End
      Begin VB.Menu M_asitosnma 
         Caption         =   "Asientos Manuales"
      End
   End
   Begin VB.Menu M_salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "inicio_CGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984




Private Sub btnsale_Click()
inicio.Show
Unload Me
End Sub




Private Sub Command1_Click()
Call nivel_acceso(7)
If para.id_grupo_modulo_actual >= 2 Then
  CGR_CUENTAS0.Show
Else
  Call sinpermisos

End If
End Sub

Private Sub Command3_Click()
Call nivel_acceso(7)
If para.id_grupo_modulo_actual >= 5 Then
 cgr_balance.Show
Else
  Call sinpermisos

End If
End Sub

Private Sub Command2_Click()
Call nivel_acceso(7)
If para.id_grupo_modulo_actual >= 5 Then
 cgr_mayores.Show
Else
  Call sinpermisos

End If
End Sub

Private Sub Command4_Click()
Call nivel_acceso(7)
If para.id_grupo_modulo_actual >= 5 Then
 ABM_periodos.Show
Else
  Call sinpermisos

End If
End Sub

Private Sub Command5_Click()
If c_periodo.ListIndex < 0 Then
  c_periodo.ListIndex = buscaindice(c_periodo, para.id_periodo_contable)
End If
J = MsgBox("Confirma Seleccionar Periodo: " & c_periodo & " para Trabajar", 4)
If J = 6 Then
   Set rs = New ADODB.Recordset
   q = "select * from g0 where [sucursal] = 0"
   rs.Open q, cn1, adOpenDynamic, adLockOptimistic
   If Not rs.EOF And Not rs.BOF Then
     rs("id_periodo_contable") = c_periodo.ItemData(c_periodo.ListIndex)
     rs.Update
   
     para.id_periodo_contable = c_periodo.ItemData(c_periodo.ListIndex)
     
End If
Call barracgr(Me)
End If

End Sub

Private Sub Command6_Click()
Call nivel_acceso(7)
If para.id_grupo_modulo_actual >= 5 Then
  abm_asientos0.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub Command7_Click()
Call nivel_acceso(7)
If para.id_grupo_modulo_actual >= 5 Then
  cgr_diario.Show
Else
  Call sinpermisos

End If
End Sub

Private Sub Form_Activate()
Call barracgr(Me)
c_periodo.ListIndex = buscaindice(c_periodo, para.id_periodo_contable)
Label7 = para.impresora_actual
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF12 Then
  gen_tools.Show
End If
End Sub

Private Sub Form_Load()

Call titulos(Me)
Call carga_periodos(c_periodo)


End Sub





Private Sub M_admasit_Click()
Call nivel_acceso(7)
If para.id_grupo_modulo_actual >= 5 Then
 cgr_admasientos.Show
End If
End Sub

Private Sub M_asa_Click()
Call nivel_acceso(7)
If para.id_grupo_modulo_actual >= 5 Then
  
  cgr_verasientos.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub M_asientos_Click()
Call nivel_acceso(7)
If para.id_grupo_modulo_actual >= 8 Then
 cgr_aGRUPA.Show
Else
  Call sinpermisos

End If
 
 
End Sub

Private Sub M_asitosnma_Click()
cgr_abmasientos_p.Show
cgr_abmasientos_p.t_funcion = "A"
End Sub

Private Sub M_maau_Click()
Call nivel_acceso(7)
If para.id_grupo_modulo_actual >= 5 Then
  cgr_vermayores.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub M_periodos_Click()
Call nivel_acceso(7)
If para.id_grupo_modulo_actual >= 5 Then
 ABM_periodos.Show
Else
  Call sinpermisos

End If
End Sub

Private Sub M_plnac_Click()
Call nivel_acceso(7)
If para.id_grupo_modulo_actual >= 5 Then
 cgr_plancuentas.Show
Else
  Call sinpermisos

End If
End Sub

Private Sub M_salir_Click()
inicio.Show
Unload Me
End Sub


