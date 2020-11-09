VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form inicio_compras 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MODULO  COMPRAS"
   ClientHeight    =   8520
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   11970
   FontTransparent =   0   'False
   Icon            =   "inicio_compras.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8520
   ScaleWidth      =   11970
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CAJA"
      Height          =   975
      Left            =   4920
      TabIndex        =   24
      Top             =   3000
      Width           =   5415
      Begin MSComctlLib.Toolbar Toolbar6 
         Height          =   645
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1138
         ButtonWidth     =   2037
         ButtonHeight    =   1032
         Appearance      =   1
         ImageList       =   "ImageList6"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B1"
               Object.ToolTipText     =   "Caja Diaria"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B2"
               Object.ToolTipText     =   "Cartera de Cheques de Tercero"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B3"
               Object.ToolTipText     =   "Informe de Resultados"
               ImageIndex      =   5
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impresora Actual del Sistema"
      Height          =   735
      Left            =   4920
      TabIndex        =   21
      Top             =   7080
      Width           =   4815
      Begin VB.CommandButton Command5 
         Height          =   375
         Left            =   4080
         Picture         =   "inicio_compras.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label7"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CONSULTAS RAPIDAS PROVEEDORES"
      Height          =   1095
      Left            =   4920
      TabIndex        =   19
      Top             =   1560
      Width           =   6495
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   645
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   1138
         ButtonWidth     =   1773
         ButtonHeight    =   1032
         Appearance      =   1
         ImageList       =   "ImageList3"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B1"
               Object.ToolTipText     =   "Estado Cuenta Proveedores"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B2"
               Object.ToolTipText     =   "Saldos Proveedores"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B3"
               Object.ToolTipText     =   "Ver Comprobantes Compra"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B4"
               Object.ToolTipText     =   "Productos en O.C."
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B5"
               Object.ToolTipText     =   "Productos Pedidos"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B6"
               Object.ToolTipText     =   "Listado Iva Compras"
               ImageIndex      =   6
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "EMISION DE COMPROBANTES PROVEEDORES"
      Height          =   1095
      Left            =   4920
      TabIndex        =   17
      Top             =   120
      Width           =   6495
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   645
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   1138
         ButtonWidth     =   1773
         ButtonHeight    =   1032
         Appearance      =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B1"
               Description     =   "Emision de Ordenes de Compra"
               Object.ToolTipText     =   "Emision de Ordenes de Compra"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B2"
               Description     =   "Recibos por Cobranzas"
               Object.ToolTipText     =   "Ingreso Comprobantes de Compra"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B3"
               Description     =   "REMITOS"
               Object.ToolTipText     =   "Emision de Ordenes de Pago"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B4"
               Object.ToolTipText     =   "Cierre Fin de Mes"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B5"
               Description     =   "Solicitud de Cotizacion"
               Object.ToolTipText     =   "Solicitud de Cotizacion"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B6"
               Object.ToolTipText     =   "Registro de Faltantes en Stock"
               ImageIndex      =   6
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ARCHIVOS MAESTROS"
      Height          =   1215
      Left            =   960
      TabIndex        =   15
      Top             =   120
      Width           =   3135
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   870
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   1535
         ButtonWidth     =   2275
         ButtonHeight    =   1429
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "PROVEEDORES"
               Key             =   "B1"
               Description     =   "Archivo de Clientes"
               Object.ToolTipText     =   "Archivo de Proveedores"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "LISTA PRECIOS"
               Key             =   "B2"
               Description     =   "Listado de Productos y Precios"
               Object.ToolTipText     =   "Lista de Productos y Precios"
               ImageIndex      =   2
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
      Top             =   5640
      Width           =   2055
      Begin VB.Image Image1 
         Height          =   480
         Left            =   720
         Picture         =   "inicio_compras.frx":0614
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "COMPRAS"
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
         Picture         =   "inicio_compras.frx":091E
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
         Picture         =   "inicio_compras.frx":11A0
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
      Top             =   8265
      Width           =   11970
      _ExtentX        =   21114
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
            TextSave        =   "02/01/2018"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "01:06 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":1A22
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":1D3C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   240
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   60
      ImageHeight     =   33
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":2056
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":27B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":2F15
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":3664
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":3D8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":4510
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
      ImageWidth      =   60
      ImageHeight     =   33
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":4C41
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":526F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":58C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":5ED5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":6534
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":6B5F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList6 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   70
      ImageHeight     =   33
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":7231
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":797E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":8107
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":88A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_compras.frx":9039
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   240
      TabIndex        =   26
      Top             =   5520
      Width           =   4575
   End
   Begin VB.Menu M_tablas 
      Caption         =   "&Tablas"
      Begin VB.Menu M_Proveedores 
         Caption         =   "Proveedores"
      End
      Begin VB.Menu M_productos 
         Caption         =   "Productos"
      End
      Begin VB.Menu M_trans 
         Caption         =   "Transportes"
         Begin VB.Menu M_abmtrasnp 
            Caption         =   "Transportes"
         End
         Begin VB.Menu M_camiones 
            Caption         =   "Camiones/Unidades"
         End
      End
      Begin VB.Menu M_perc 
         Caption         =   "Percepciones"
      End
   End
   Begin VB.Menu M_consultas 
      Caption         =   "Consultas"
      Begin VB.Menu M_estadocta 
         Caption         =   "Estado de Cuenta Proveedores"
      End
      Begin VB.Menu M_vercomp 
         Caption         =   "Comprobantes Ingresados"
      End
      Begin VB.Menu M_saldos 
         Caption         =   "Saldos Proveedores"
      End
      Begin VB.Menu M_subdiario 
         Caption         =   "Subdiario de Compras"
      End
      Begin VB.Menu M_oc 
         Caption         =   "Ordenes de Compra"
         Begin VB.Menu M_prodoc 
            Caption         =   "Productos en O.C."
         End
         Begin VB.Menu M_emitidas 
            Caption         =   "O.C. Emitidas"
         End
      End
      Begin VB.Menu M_proding 
         Caption         =   "Productos ingresados "
      End
      Begin VB.Menu M_imoos 
         Caption         =   "Consultas Impositivas"
         Begin VB.Menu m_listaiva 
            Caption         =   "Listado de Iva Compras"
         End
         Begin VB.Menu M_liret 
            Caption         =   "Listado de Retenciones y Percepciones Realizadas"
         End
         Begin VB.Menu M_retrec 
            Caption         =   "Listado de Retenciones y Percepciones Recibidas"
         End
         Begin VB.Menu M_conret 
            Caption         =   "Calculo de Retenciones"
         End
         Begin VB.Menu M_ley23966 
            Caption         =   "Subsidio ley 23966 Art. 15"
         End
         Begin VB.Menu M_posicioniva 
            Caption         =   "Pisicion frente al IVA"
         End
         Begin VB.Menu M_verfapo 
            Caption         =   "Verifica Comprobantes Apócrifos"
         End
      End
      Begin VB.Menu m_HISTORICO 
         Caption         =   "Historico de compras por producto"
      End
      Begin VB.Menu M_vto 
         Caption         =   "Vencimientos"
      End
   End
   Begin VB.Menu M_uitl 
      Caption         =   "Utiles"
      Begin VB.Menu M_configura 
         Caption         =   "Configurar Comprobantes"
      End
      Begin VB.Menu M_tools 
         Caption         =   "Utilitarios Varios(Tools)"
      End
      Begin VB.Menu M_corrigecuit 
         Caption         =   "Corrige Cuits"
      End
      Begin VB.Menu M_cierra 
         Caption         =   "Cierra mes"
      End
      Begin VB.Menu M_factursa 
         Caption         =   "Faturas Apocrifas "
         Begin VB.Menu M_actuapo 
            Caption         =   "Actualizar padron facturas apocrifas"
         End
         Begin VB.Menu M_consultcuit 
            Caption         =   "Consulta Individual por CUIT"
         End
         Begin VB.Menu M_cargacomp 
            Caption         =   "Verificar comprobantes desde Excel"
         End
      End
      Begin VB.Menu M_actuch 
         Caption         =   "Actualiza Valor de Ch. terc. a la cotiz. actual"
      End
   End
   Begin VB.Menu M_salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "inicio_compras"
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
gen_seleccionarimp.Show
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
Label7 = para.impresora_actual
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF12 Then
  gen_tools.Show
End If
End Sub

Private Sub Form_Load()
Call titulos(Me)
Set rs = New ADODB.Recordset
q = "select * from g0 where [sucursal] = 0"
rs.Open q, cn1
io = rs("id_obraactual")
Set rs = Nothing
Call activaobra(io)


Select Case para.HABILITACION
Case Is = 1825, Is = 9999  'facturacion fletes
     M_ley23966.Visible = True
Case Else
     M_ley23966.Visible = False
End Select



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

Private Sub M_abmtrasnp_Click()
ABM_PROv.Show
End Sub

Private Sub M_actuapo_Click()
gen_factapocrifas.Show
End Sub

Private Sub M_actuch_Click()
cyb_actucot_ch_terc.Show
End Sub

Private Sub M_camiones_Click()
GEN_ABMCAMION.Show
End Sub

Private Sub M_cargacomp_Click()
con_busca_comp_apoc_excel.Show
End Sub

Private Sub M_cierra_Click()
 Call nivel_acceso(2)
    If para.id_grupo_modulo_actual >= 9 Then
      com_cierremes.Show
    Else
      Call sinpermisos
    End If
End Sub

Private Sub M_configura_Click()
Call nivel_acceso(2)
If para.id_grupo_modulo_actual >= 9 Then
  com_config_comp.Show
Else
  Call sinpermisos
End If

End Sub

Private Sub M_conret_Click()
calcula_ret.Show
End Sub

Private Sub M_consultcuit_Click()
com_consultaapoc.Show
End Sub

Private Sub M_corrigecuit_Click()
Call verificacuits
End Sub
Sub verificacuits()
h = MsgBox("Formatea(numerico sin guines de longitud 11) los cuit de los Clientes   . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from a1"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    If rs("cod_tipoiva") <> 3 And rs("cod_tipoiva") <> 8 Then
       'lleva cuit
       If Len(rs("cuit")) = 13 Then
          c = Mid$(rs("cuit"), 1, 2) & Mid$(rs("cuit"), 4, 8) & Mid$(rs("cuit"), 13, 1)
          rs("cuit") = c
          rs.Update
       End If
     Else
        If Val(rs("cuit")) <= 0 Then
         rs("cuit") = 0
        End If
     
     End If
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If
End Sub

Private Sub M_emitidas_Click()
con_veroc.Show
End Sub

Private Sub M_estadocta_Click()
Call nivel_acceso(2)
If para.id_grupo_modulo_actual >= 4 Then
  con_estadocuenta.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub m_HISTORICO_Click()
Call nivel_acceso(2)
If para.id_grupo_modulo_actual >= 5 Then
  con_HISTORICOcompras.Show
  con_HISTORICOcompras.Check1 = 1
Else
  Call sinpermisos
End If

End Sub

Private Sub M_ley23966_Click()
con_ley23966.Show
End Sub

Private Sub M_liret_Click()
con_retperc.Show
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

Private Sub M_posicioniva_Click()
gen_posicioniva.Show
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

Private Sub M_retrec_Click()
vta_retyperc.Show
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

Private Sub M_tools_Click()
gen_tools.Show
End Sub

Private Sub M_vercomp_Click()
Call nivel_acceso(2)
If para.id_grupo_modulo_actual >= 4 Then
  con_vercomp.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub M_verfapo_Click()
con_busca_comp_apoc.Show
End Sub

Private Sub M_vto_Click()
com_vencimientos.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case Is = "B1"
    ABM_PROv.Show

  Case Is = "B2"
    vta_listaprecios.Show

End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case Is = "B1"
    Call nivel_acceso(2)
    If para.id_grupo_modulo_actual >= 5 Then
     ABM_OC.Show
    Else
     Call sinpermisos
    End If
  Case Is = "B2"
    Call nivel_acceso(2)
    If para.id_grupo_modulo_actual >= 4 Then
      ABM_COMP_COMPRA.Show
      ABM_COMP_COMPRA.t_funcion = "D"
    Else
      Call sinpermisos
    End If
  Case Is = "B3"
   Call nivel_acceso(2)
   If para.id_grupo_modulo_actual >= 7 Then
      op.Show
   Else
      Call sinpermisos
   End If
  Case Is = "B4"
    Call nivel_acceso(2)
    If para.id_grupo_modulo_actual >= 9 Then
      com_cierremes.Show
    Else
      Call sinpermisos
    End If
Case Is = "B5"
    Call nivel_acceso(2)
    If para.id_grupo_modulo_actual >= 5 Then
      ABM_cotizacion.Show
    Else
      Call sinpermisos
    End If
Case Is = "B6"
    Call nivel_acceso(2)
    If para.id_grupo_modulo_actual >= 4 Then
      com_faltantes.Show
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
     con_estadocuenta.Show
    Else
     Call sinpermisos
    End If
 Case Is = "B2"
    Call nivel_acceso(2)
    If para.id_grupo_modulo_actual >= 4 Then
      con_saldosprov.Show
    Else
      Call sinpermisos
    End If
  Case Is = "B3"
   Call nivel_acceso(2)
   If para.id_grupo_modulo_actual >= 4 Then
     con_vercomp.Show
   Else
    Call sinpermisos
   End If
Case Is = "B4"
    Call nivel_acceso(2)
    If para.id_grupo_modulo_actual >= 4 Then
      ver_PROD_oc.Show
    Else
      Call sinpermisos
    End If
Case Is = "B5"
  Call nivel_acceso(2)
  If para.id_grupo_modulo_actual >= 2 Then
    abm_solmat.Show
  Else
    Call sinpermisos
  End If
Case Is = "B6"
  Call nivel_acceso(2)
  If para.id_grupo_modulo_actual >= 5 Then
    con_ivacompras.Show
  Else
    Call sinpermisos
  End If

End Select

End Sub

Private Sub Toolbar6_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case Is = "B1"
    Call nivel_acceso(3)
    If para.id_grupo_modulo_actual >= 4 Then
     cja_cajadiaria.Show
    Else
     Call sinpermisos
    End If
 Case Is = "B2"
   Call nivel_acceso(3)
   If para.id_grupo_modulo_actual >= 4 Then
     cyb_carterach.Show
   Else
    Call sinpermisos
   End If
 Case Is = "B3"
   Call nivel_acceso(1)
   If para.id_grupo_modulo_actual >= 7 Then
     vta_gerencial1.Show
   Else
    Call sinpermisos
   End If


End Select

End Sub
