VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form inicio_produccion 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MODULO PRODUCCION"
   ClientHeight    =   8355
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   12405
   FontTransparent =   0   'False
   Icon            =   "inicio_produccion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8355
   ScaleWidth      =   12405
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Costos - Estructura Productos"
      Height          =   2055
      Left            =   240
      TabIndex        =   26
      Top             =   2880
      Width           =   3375
      Begin VB.CommandButton Command9 
         Caption         =   "ABM Piezas "
         Height          =   495
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Estructura de Productos"
         Height          =   495
         Left            =   240
         TabIndex        =   28
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Planilla de Costos"
         Height          =   495
         Left            =   240
         TabIndex        =   27
         Top             =   1440
         Width           =   2895
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impresora Actual del Sistema"
      Height          =   735
      Left            =   4920
      TabIndex        =   23
      Top             =   7080
      Width           =   4815
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   4080
         Picture         =   "inicio_produccion.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label7"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modulo "
      Height          =   1455
      Left            =   6120
      TabIndex        =   20
      Top             =   5640
      Width           =   2055
      Begin VB.Image Image1 
         Height          =   480
         Left            =   720
         Picture         =   "inicio_produccion.frx":0614
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PRODUCCION"
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
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Consultas"
      Height          =   2655
      Left            =   8400
      TabIndex        =   17
      Top             =   360
      Width           =   3375
      Begin VB.CommandButton Command12 
         Caption         =   "Ver Ordenes de Empaque"
         Height          =   495
         Left            =   240
         TabIndex        =   32
         Top             =   1800
         Width           =   2895
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Ver Comprobantes Produccion"
         Height          =   495
         Left            =   240
         TabIndex        =   31
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Materiales Pedidos"
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Produccion"
      Height          =   2055
      Left            =   240
      TabIndex        =   14
      Top             =   360
      Width           =   3375
      Begin VB.CommandButton Command5 
         Caption         =   "Presupuesto Obra"
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   2895
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ABM Obras o Trabajos"
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pedidos de Materias Prima y Productos"
      Height          =   2655
      Left            =   4320
      TabIndex        =   13
      Top             =   360
      Width           =   3375
      Begin VB.CommandButton Command10 
         Caption         =   "Orden de Empaque"
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   1800
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Validar Estados de Solicitudes"
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Solicitud de Materiales"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   2895
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
         Picture         =   "inicio_produccion.frx":0EDE
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
         Picture         =   "inicio_produccion.frx":1760
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
      Top             =   8100
      Width           =   12405
      _ExtentX        =   21881
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
            TextSave        =   "27/02/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:44"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Menu M_consultas 
      Caption         =   "Consultas"
      Begin VB.Menu M_comping 
         Caption         =   "Ver Comprobantes Ingresados"
      End
      Begin VB.Menu M_mov_prod 
         Caption         =   "Movimientos por Productos"
      End
   End
   Begin VB.Menu M_salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "inicio_produccion"
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
prod_solmat.Show
End Sub

Private Sub Command10_Click()
pro_empaque.Show

End Sub

Private Sub Command11_Click()
prod_vercomp.Show
End Sub

Private Sub Command12_Click()
prod_verempaque.Show
End Sub

Private Sub Command13_Click()
Set rs = New Recordset
q = "select * from pro_04 where [num_referencia] < 6000"
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
While Not rs.EOF
  rs("total_oc") = rs("total_pedido")
  rs("total_recibido") = rs("total_pedido")
  
  rs.MoveNext
Wend
Set rs = Nothing
MsgBox ("Termino")
  

End Sub

Private Sub Command2_Click()
abm_solmat.Show
End Sub

Private Sub Command3_Click()
J = MsgBox("Confirma estado", 4)
If J = 6 Then
  espere.Show
  espere.Refresh
  espere.ProgressBar1.Min = 1
  espere.ProgressBar1.Max = 4
  
  Set rs = New ADODB.Recordset
  q = "select * from pro_04"
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  While Not rs.EOF
    'validar total pedido
    espere.ProgressBar1.Value = 1
    Set rs2 = New ADODB.Recordset
    q = "select * from pro_05 where [num_referencia] = " & rs("num_referencia") & " and [tipo_comprobante] = 1"
    rs2.Open q, cn1
    tp = 0
    While Not rs2.EOF
      tp = tp + rs2("cantidad")
      rs2.MoveNext
    Wend
    Set rs2 = Nothing
    
    espere.ProgressBar1.Value = 2
    'validar total en oc
    Set rs2 = New ADODB.Recordset
    q = "select * from pro_05 where [num_referencia] = " & rs("num_referencia") & " and [tipo_comprobante] = 2"
    rs2.Open q, cn1
    toc = 0
    While Not rs2.EOF
      toc = toc + rs2("cantidad")
      rs2.MoveNext
    Wend
    Set rs2 = Nothing
    
    espere.ProgressBar1.Value = 3
    'validar total recibido
    Set rs2 = New ADODB.Recordset
    q = "select * from pro_05 where [num_referencia] = " & rs("num_referencia") & " and [tipo_comprobante] = 3"
    rs2.Open q, cn1
    tr = 0
    While Not rs2.EOF
      tr = tr + rs2("cantidad")
      rs2.MoveNext
    Wend
    Set rs2 = Nothing
  
  
    espere.ProgressBar1.Value = 4
    If tp > toc Then
      ep = "I"
    Else
      ep = "C"
    End If
    
    If toc > tr Then
      eoc = "I"
    Else
      eoc = "C"
    End If
    
    rs("total_oc") = toc
    rs("total_pedido") = tp
    rs("total_recibido") = tr
    rs("estado_pedido") = ep
    rs("estado_oc") = eoc
    rs.Update
    
    rs.MoveNext
  Wend
  Set rs = Nothing
    
  
  
  Unload espere
End If
End Sub

Private Sub Command4_Click()
ABM_OBRAS.Show
End Sub


Private Sub Command6_Click()
gen_seleccionarimp.Show
End Sub

Private Sub Command7_Click()
pro_costos.Show
End Sub

Private Sub Command8_Click()
pro_estructura.Show
End Sub

Private Sub Command9_Click()
pro_ABM_pieza.Show
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





Private Sub M_comping_Click()
prod_vercomp.Show

End Sub

Private Sub M_mov_prod_Click()
stk_movprod.Show
End Sub


Private Sub M_salir_Click()
inicio.Show
Unload Me
End Sub


Private Sub M_seguirped_Click()

End Sub
