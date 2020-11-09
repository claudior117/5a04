VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form inicio_caja 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MODULO CAJA"
   ClientHeight    =   8655
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   12570
   FontTransparent =   0   'False
   Icon            =   "inicio_caja.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8655
   ScaleWidth      =   12570
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Operaciones Varias"
      Height          =   975
      Left            =   3720
      TabIndex        =   24
      Top             =   1920
      Width           =   7815
      Begin VB.CommandButton Command4 
         Caption         =   "Transferencia Interna entre Cajas(TICC)"
         Height          =   495
         Left            =   4200
         TabIndex        =   27
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Informe Cierre"
         Height          =   495
         Left            =   2160
         TabIndex        =   26
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Transferencias a Banco"
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   360
      TabIndex        =   22
      Top             =   1680
      Width           =   2055
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   120
         Picture         =   "inicio_caja.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impresora Actual del Sistema"
      Height          =   735
      Left            =   4920
      TabIndex        =   19
      Top             =   7080
      Width           =   4815
      Begin VB.CommandButton Command5 
         Height          =   375
         Left            =   4080
         Picture         =   "inicio_caja.frx":0C18
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label7"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ARCHIVOS MAESTROS"
      Height          =   1215
      Left            =   360
      TabIndex        =   17
      Top             =   360
      Width           =   2055
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   870
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   1535
         ButtonWidth     =   2434
         ButtonHeight    =   1429
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Formas de Pago"
               Key             =   "B2"
               Description     =   "Listado de Productos y Precios"
               Object.ToolTipText     =   "Lista de Productos y Precios"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CAJA DIARIA"
      Height          =   1095
      Left            =   3720
      TabIndex        =   15
      Top             =   360
      Width           =   7815
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   645
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1138
         ButtonWidth     =   2037
         ButtonHeight    =   1032
         Appearance      =   1
         ImageList       =   "ImageList3"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B1"
               Object.ToolTipText     =   "Caja Diaria"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B2"
               Object.ToolTipText     =   "Composicion de los saldos de caja"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B3"
               Object.ToolTipText     =   "Informe General de Caja"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B4"
               Object.ToolTipText     =   "Cartera de Cheques de Tercero"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B5"
               Object.ToolTipText     =   "Resultados"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B6"
               Object.ToolTipText     =   "Cierre de Caja"
               ImageIndex      =   5
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modulo "
      Height          =   1455
      Left            =   6360
      TabIndex        =   13
      Top             =   5640
      Width           =   2055
      Begin VB.Image Image1 
         Height          =   480
         Left            =   720
         Picture         =   "inicio_caja.frx":0F22
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "CAJA"
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
         Picture         =   "inicio_caja.frx":122C
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
         Picture         =   "inicio_caja.frx":1AAE
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
      Top             =   8400
      Width           =   12570
      _ExtentX        =   22172
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
      Left            =   1200
      Top             =   2280
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
            Picture         =   "inicio_caja.frx":2330
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_caja.frx":264A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   0
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   70
      ImageHeight     =   33
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_caja.frx":2964
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_caja.frx":30B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_caja.frx":383A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_caja.frx":3FD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_caja.frx":476C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_caja.frx":5C28
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
      TabIndex        =   28
      Top             =   5520
      Width           =   4575
   End
   Begin VB.Menu M_tablas 
      Caption         =   "&Tablas"
      Begin VB.Menu M_fp 
         Caption         =   "Formas de Pago"
      End
   End
   Begin VB.Menu M_consultas 
      Caption         =   "&Consultas"
      Begin VB.Menu M_comparativo 
         Caption         =   "Cash Flow"
      End
      Begin VB.Menu M_igchtç 
         Caption         =   "Inf. Auditoria(Ing y Egr Ch terc a caja)"
      End
   End
   Begin VB.Menu M_salir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "inicio_caja"
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
Call nivel_acceso(3)
If para.id_grupo_modulo_actual >= 8 Then
 CGR_CUENTAS0.Show
Else
 Call sinpermisos
End If
End Sub


Private Sub Command2_Click()
Call nivel_acceso(3)
If para.id_grupo_modulo_actual >= 8 Then
  cja_transf_banco.Show
End If
End Sub

Private Sub Command6_Click()
Call nivel_acceso(3)
If para.id_grupo_modulo_actual >= 5 Then
  cyb_cajadiaria.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub Command3_Click()
cja_informecierre.Show
End Sub

Private Sub Command4_Click()
Call nivel_acceso(3)
If para.id_grupo_modulo_actual >= 9 Then
  cja_transf_caja.Show
End If
End Sub

Private Sub Command5_Click()
gen_seleccionarimp.Show
End Sub

Private Sub Form_Activate()
Call barraesag(Me)
Label7 = para.impresora_actual
End Sub

Private Sub Form_Load()
Call titulos(Me)

Exit Sub

e1:
  MsgBox ("Error al Inicializar Parametros INICIO.LOAD")
  End

End Sub




Private Sub M_comparativo_Click()
   Call nivel_acceso(3)
   If para.id_grupo_modulo_actual > 4 Then
     cja_detallemov2.Show
   Else
     Call sinpermisos
   End If
End Sub

Private Sub M_fp_Click()
cyb_ABM_FP.Show
End Sub

Private Sub M_informeCaja_Click()

End Sub

Private Sub M_informeresultado_Click()

End Sub

Private Sub M_igchtç_Click()
cyb_carterach2.Show
End Sub

Private Sub M_salir_Click()
inicio.Show
Unload Me
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case Is = "B2"
   Call nivel_acceso(3)
   If para.id_grupo_modulo_actual > 1 Then
     cyb_ABM_FP.Show
   End If
End Select

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case Is = "B1"
    Call nivel_acceso(3)
    If para.id_grupo_modulo_actual >= 5 Then
     cja_cajadiaria.Show
    Else
     Call sinpermisos
    End If
 Case Is = "B2"
    Call nivel_acceso(3)
    If para.id_grupo_modulo_actual >= 5 Then
      cyb_cajadiaria.Show
    Else
      Call sinpermisos
    End If
 Case Is = "B3"
   Call nivel_acceso(3)
   If para.id_grupo_modulo_actual >= 2 Then
     cja_detallemov.Show
   Else
    Call sinpermisos
   End If
 Case Is = "B4"
   Call nivel_acceso(3)
   If para.id_grupo_modulo_actual >= 5 Then
     cyb_carterach.Show
   Else
    Call sinpermisos
   End If
 Case Is = "B5"
   Call nivel_acceso(3)
   If para.id_grupo_modulo_actual >= 8 Then
     vta_gerencial1.Show
   Else
    Call sinpermisos
   End If

 Case Is = "B6"
   Call nivel_acceso(3)
   If para.id_grupo_modulo_actual >= 8 Then
     cja_cierremes.Show
   Else
    Call sinpermisos
   End If


End Select

End Sub
