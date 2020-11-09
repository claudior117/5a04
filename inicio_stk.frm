VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form inicio_stk 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MODULO STOCK y CONTROL DE MATERIALES"
   ClientHeight    =   8400
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   12330
   FontTransparent =   0   'False
   Icon            =   "inicio_stk.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8400
   ScaleWidth      =   12330
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "MOVIMIENTOS DE STOCK"
      Height          =   1095
      Left            =   3720
      TabIndex        =   17
      Top             =   600
      Width           =   4695
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   645
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1138
         ButtonWidth     =   2566
         ButtonHeight    =   1032
         Appearance      =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B1"
               Description     =   "Entradas de mercaderia"
               Object.ToolTipText     =   "Entradas de Mercaderias"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B2"
               Description     =   "Salidas mercaderia"
               Object.ToolTipText     =   "Salidas de Mercaderias"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B3"
               Description     =   "Ajustes de Stock"
               Object.ToolTipText     =   "Ajustes de Stock"
               ImageIndex      =   1
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ARCHIVOS MAESTROS"
      Height          =   1215
      Left            =   240
      TabIndex        =   15
      Top             =   600
      Width           =   3135
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   870
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   1535
         ButtonWidth     =   2275
         ButtonHeight    =   1429
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "PRODUCTOS"
               Key             =   "B1"
               Description     =   "Archivo de Clientes"
               Object.ToolTipText     =   "Archivo de Proveedores"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "LISTA PRECIOS"
               Key             =   "B2"
               Description     =   "Listado de Productos y Precios"
               Object.ToolTipText     =   "Lista de Productos y Precios"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modulo "
      Height          =   1455
      Left            =   6120
      TabIndex        =   13
      Top             =   6360
      Width           =   2055
      Begin VB.Image Image1 
         Height          =   480
         Left            =   720
         Picture         =   "inicio_stk.frx":030A
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "STOCK"
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
         Left            =   480
         TabIndex        =   14
         Top             =   960
         Width           =   975
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
         Picture         =   "inicio_stk.frx":0614
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
         Picture         =   "inicio_stk.frx":0E96
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
      Width           =   12330
      _ExtentX        =   21749
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
            TextSave        =   "06/08/2017"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "05:57 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_stk.frx":1718
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_stk.frx":1A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_stk.frx":1DA6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1200
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   90
      ImageHeight     =   33
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_stk.frx":2062
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_stk.frx":2741
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_stk.frx":2FC3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu M_tablas 
      Caption         =   "&Tablas"
      Begin VB.Menu M_grupos 
         Caption         =   "Grupos"
      End
      Begin VB.Menu M_deptos 
         Caption         =   "Departamentos"
      End
      Begin VB.Menu M_marcas 
         Caption         =   "Marcas"
      End
   End
   Begin VB.Menu M_consultas 
      Caption         =   "Consultas"
      Begin VB.Menu M_mov_prod 
         Caption         =   "Movimientos por Productos"
      End
      Begin VB.Menu M_movf 
         Caption         =   "Movimientos por Fecha"
      End
      Begin VB.Menu M_comping 
         Caption         =   "Comrpobantes Ingresados"
      End
      Begin VB.Menu M_inventario 
         Caption         =   "Inventario"
      End
   End
   Begin VB.Menu M_utiles 
      Caption         =   "Utiles"
      Begin VB.Menu M_reindexastock 
         Caption         =   "Actualiza Stock Instantaneo desde Movimientos"
      End
      Begin VB.Menu M_ajusta 
         Caption         =   "Ajusta stock Movimientos desde Stock Instantaneo"
      End
   End
   Begin VB.Menu M_salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "inicio_stk"
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
ABM_PROD.Show
End Sub

Private Sub Command2_Click()
stk_movint.Show
End Sub

Private Sub Command3_Click()
abm_solmat.Show
End Sub

Private Sub Command4_Click()
ABM_OBRAS.Show
End Sub

Private Sub Command5_Click()
ver_PROD_oc.Show
End Sub

Private Sub Command6_Click()
stk_seguirpedidos.Show
End Sub

Private Sub Form_Activate()
Call barraesag(Me)

End Sub

Private Sub Form_Load()
Call titulos(Me)

Exit Sub

e1:
  MsgBox ("Error al Inicializar Parametros INICIO.LOAD")
  End

End Sub




Private Sub M_ajusta_Click()
stk_ajustedesdeinst.Show
End Sub

Private Sub M_comping_Click()
stk_vercomp.Show
End Sub

Private Sub M_deptos_Click()
ABM_deptoS.Show
End Sub

Private Sub M_grupos_Click()
ABM_grupos.Show
End Sub

Private Sub M_inventario_Click()
stk_inventario.Show
End Sub

Private Sub M_marcas_Click()
ABM_marcas.Show
End Sub

Private Sub M_mov_prod_Click()
stk_movprod.Show
End Sub

Private Sub M_obras_Click()
ABM_OBRAS.Show
End Sub

Private Sub M_prod_Click()
ABM_PROD.Show
End Sub

Private Sub M_movf_Click()
stk_movprod2.Show
End Sub

Private Sub M_reindexastock_Click()
Set cl_stock = New STOCK
cl_stock.actualizastock
Set cl_stock = Nothing

End Sub

Private Sub M_salir_Click()
inicio.Show
Unload Me
End Sub


Private Sub M_seguirped_Click()
stk_seguirpedidos.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case Is = "B1"
    ABM_PROD.Show

  Case Is = "B2"
    vta_listaprecios.Show

End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case Is = "B1"
    stk_ingreso.Show

  Case Is = "B2"
     stk_egreso.Show
     
  Case Is = "B3"
  stk_movint.Show

    

End Select
End Sub
