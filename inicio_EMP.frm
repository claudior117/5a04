VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form inicio_empleados 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MODULO EMPLEADOS"
   ClientHeight    =   8400
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   11835
   FontTransparent =   0   'False
   Icon            =   "inicio_EMP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8400
   ScaleWidth      =   11835
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CONSULTAS RAPIDAS "
      Height          =   1095
      Left            =   4920
      TabIndex        =   18
      Top             =   1560
      Width           =   5415
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   780
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1376
         ButtonWidth     =   2566
         ButtonHeight    =   1270
         Appearance      =   1
         ImageList       =   "ImageList3"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B1"
               Object.ToolTipText     =   "Estado Cuenta Empleados"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B2"
               Object.ToolTipText     =   "Saldos Cuenta Empleados"
               ImageIndex      =   2
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "REGISTRO"
      Height          =   1095
      Left            =   4920
      TabIndex        =   16
      Top             =   120
      Width           =   3855
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   780
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1376
         ButtonWidth     =   2566
         ButtonHeight    =   1270
         Appearance      =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B1"
               Description     =   "Registra Ingresos y Egresos cuenta Empleados "
               Object.ToolTipText     =   "Registra Ingresos y Egresos cuenta Empleados"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B2"
               ImageIndex      =   2
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ARCHIVOS MAESTROS"
      Height          =   1455
      Left            =   960
      TabIndex        =   15
      Top             =   120
      Width           =   2055
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   1095
         Left            =   480
         TabIndex        =   20
         Top             =   240
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   1931
         ButtonWidth     =   1958
         ButtonHeight    =   1826
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "EMPLEADOS"
               Key             =   "B1"
               Description     =   "Archivo de Clientes"
               Object.ToolTipText     =   "Archivo de Proveedores"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modulo "
      Height          =   1455
      Left            =   6240
      TabIndex        =   13
      Top             =   6360
      Width           =   2055
      Begin VB.Image Image1 
         Height          =   705
         Left            =   720
         Picture         =   "inicio_EMP.frx":030A
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "EMPLEADOS"
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
         Left            =   360
         TabIndex        =   14
         Top             =   960
         Width           =   1335
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
         Picture         =   "inicio_EMP.frx":088F
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
         Picture         =   "inicio_EMP.frx":1111
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
      Top             =   8145
      Width           =   11835
      _ExtentX        =   20876
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
            TextSave        =   "03/02/2020"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "17:44"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   41
      ImageHeight     =   47
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_EMP.frx":1993
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   120
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   90
      ImageHeight     =   42
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_EMP.frx":1F28
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_EMP.frx":223E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   120
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   90
      ImageHeight     =   42
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_EMP.frx":274A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_EMP.frx":29FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu M_tablas 
      Caption         =   "&Tablas"
      Begin VB.Menu M_Proveedores 
         Caption         =   "Empleados"
      End
   End
   Begin VB.Menu M_consultas 
      Caption         =   "Consultas"
   End
   Begin VB.Menu M_salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "inicio_empleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984




Private Sub btnsale_Click()
inicio.Show
Unload Me
End Sub



Private Sub CommandButton1_Click()
inicio_campaña.Show
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
Call nivel_acceso(2)
If para.id_grupo_modulo_actual >= 4 Then
   ABM_COMP_COMPRA.Show
   ABM_COMP_COMPRA.t_funcion = "D"
Else
  Call sinpermisos
End If
End Sub

Private Sub Command3_Click()
Call nivel_acceso(2)
If para.id_grupo_modulo_actual >= 1 Then
  con_vercomp.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub Command4_Click()
Call nivel_acceso(2)
If para.id_grupo_modulo_actual >= 8 Then
   op.Show
Else
  Call sinpermisos
End If

End Sub

Private Sub Command5_Click()
ver_PROD_oc.Show
End Sub

Private Sub Command6_Click()
Call nivel_acceso(2)
If para.id_grupo_modulo_actual >= 9 Then
   com_cierremes.Show
Else
  Call sinpermisos
End If

End Sub

Private Sub Command7_Click()
Call nivel_acceso(2)
If para.id_grupo_modulo_actual >= 2 Then
  abm_solmat.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub Command8_Click()
Call nivel_acceso(2)
If para.id_grupo_modulo_actual >= 2 Then
  con_estadocuenta.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub Form_Activate()
Call barraesag(Me)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF12 Then
  gen_tools.Show
End If
End Sub

Private Sub Form_Load()
Call titulos(Me)
Set rs = New adodb.Recordset
q = "select * from g0 where [sucursal] = 0"
rs.Open q, cn1
io = rs("id_obraactual")
Set rs = Nothing
Call activaobra(io)

Exit Sub

e1:
  MsgBox ("Error al Inicializar Parametros INICIO.LOAD")
  End

End Sub


Private Sub M_actualizadatos_Click()
J = InputBox$("Ingrese Password")
prueba = "N"
If J = "0969" Then
   

  MsgBox ("Proceso Terminado")

End If
End Sub




Private Sub M_obras_Click()
ABM_OBRAS.Show
End Sub

Private Sub M_conret_Click()
calcula_ret.Show
End Sub

Private Sub M_estadocta_Click()
Call nivel_acceso(2)
If para.id_grupo_modulo_actual >= 4 Then
  con_estadocuenta.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub m_listaiva_Click()
Call nivel_acceso(2)
If para.id_grupo_modulo_actual >= 5 Then
  con_ivacompras.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub M_perc_Click()
ABM_perc.Show

End Sub

Private Sub M_proding_Click()
Call nivel_acceso(2)
If para.id_grupo_modulo_actual >= 2 Then
  con_verprod.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub M_prodoc_Click()
Call nivel_acceso(2)
If para.id_grupo_modulo_actual >= 2 Then
  ver_PROD_oc.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub M_productos_Click()
ABM_PROD.Show
End Sub

Private Sub M_Proveedores_Click()
ABM_PROv.Show
End Sub

Private Sub M_saldos_Click()
Call nivel_acceso(2)
If para.id_grupo_modulo_actual >= 4 Then
     con_saldosprov.Show
Else
    Call sinpermisos
End If

End Sub

Private Sub M_salir_Click()
inicio.Show
Unload Me
End Sub


Private Sub m_v1_Click()
frmAbout.Show
End Sub

Private Sub M_subdiario_Click()
Call nivel_acceso(2)
If para.id_grupo_modulo_actual >= 4 Then
  con_subdiarioc.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub M_vercomp_Click()
Call nivel_acceso(2)
If para.id_grupo_modulo_actual >= 4 Then
  con_vercomp.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case Is = "B1"
    emp_ABM_emp.Show

End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case Is = "B1"
    Call nivel_acceso(2)
    If para.id_grupo_modulo_actual >= 5 Then
     emp_emitemov.Show
    Else
     Call sinpermisos
    End If
  
 Case Is = "B2"
    Call nivel_acceso(2)
    If para.id_grupo_modulo_actual >= 5 Then
     emp_emitegastos.Show
    Else
     Call sinpermisos
    End If

End Select
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case Is = "B1"
    Call nivel_acceso(2)
    If para.id_grupo_modulo_actual >= 4 Then
     emp_estadocuenta.Show
    Else
     Call sinpermisos
    End If
  Case Is = "B2"
    Call nivel_acceso(2)
    If para.id_grupo_modulo_actual >= 5 Then
     emp_saldos.Show
    Else
     Call sinpermisos
    End If
End Select

End Sub
