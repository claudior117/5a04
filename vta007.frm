VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0A6BE9FC-5039-11D5-98EC-0800460222F0}#1.0#0"; "IFEpson.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_recibo 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RECIBO"
   ClientHeight    =   8415
   ClientLeft      =   330
   ClientTop       =   705
   ClientWidth     =   11955
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8415
   ScaleWidth      =   11955
   Begin VB.Frame Frame8 
      Caption         =   "Saldo"
      Height          =   855
      Left            =   9240
      TabIndex        =   49
      Top             =   1440
      Width           =   2415
      Begin VB.TextBox t_saldo21 
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1335
      Left            =   9240
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox t_numintnc 
         Height          =   285
         Left            =   120
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox t_totalnc 
         Height          =   285
         Left            =   1560
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox t_ivanc 
         Height          =   285
         Left            =   1560
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox t_subtotalnc 
         Height          =   285
         Left            =   1560
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox t_numnc 
         Height          =   285
         Left            =   120
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox t_sucnc 
         Height          =   285
         Left            =   120
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   5160
      TabIndex        =   39
      Top             =   720
      Width           =   3855
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Solo Cuenta U$s"
         Height          =   315
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   1935
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Solo Cuenta $"
         Height          =   315
         Left            =   2160
         TabIndex        =   40
         Top             =   120
         Width           =   1575
      End
   End
   Begin EPSON_Impresora_Fiscal.PrinterFiscal epson1 
      Left            =   10920
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10320
      TabIndex        =   33
      Top             =   6960
      Width           =   1575
      Begin VB.CommandButton confirma 
         Height          =   615
         Left            =   120
         Picture         =   "vta007.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton salir 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta007.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   1695
      Left            =   120
      TabIndex        =   25
      Top             =   6240
      Width           =   9495
      Begin VB.TextBox t_retenciones 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6000
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox fdolar 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5280
         MaxLength       =   8
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Detalle 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         MaxLength       =   99
         TabIndex        =   11
         Top             =   1200
         Width           =   7335
      End
      Begin VB.TextBox t_pago 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox total 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         MaxLength       =   12
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox t_totald 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8040
         MaxLength       =   8
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Total Retenciones Aplicadas"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4080
         TabIndex        =   36
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Cotiz. U$s:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4080
         TabIndex        =   30
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "En concepto  de:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Total Comprobantes Aplicados"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Total a ingresar en cuenta cliente"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   8
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Total U$s:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   6720
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   9015
      Begin VB.ComboBox c_sucursal 
         Height          =   315
         ItemData        =   "vta007.frx":1104
         Left            =   7200
         List            =   "vta007.frx":1106
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox t_numop 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3120
         MaxLength       =   8
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox t_fecha 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox sucursal 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Punto Venta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5640
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Nro Recibo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comprobantes a Aplicar"
      Height          =   1695
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   11655
      Begin VB.CommandButton Command2 
         Caption         =   "Facturar"
         Height          =   195
         Left            =   10560
         TabIndex        =   38
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ret."
         Height          =   195
         Left            =   10560
         TabIndex        =   37
         Top             =   720
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   1335
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   2355
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Forma de Pago"
      Height          =   2175
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   11655
      Begin VB.TextBox t_diferencia 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10320
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   19
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox t_ingresado 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10320
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid msf2 
         Height          =   1815
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3201
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         Caption         =   "Diferencia"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   10320
         TabIndex        =   20
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Total Ing."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   10320
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cliente"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   9015
      Begin VB.CommandButton Command5 
         Height          =   255
         Left            =   8520
         Picture         =   "vta007.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   255
      End
      Begin VB.ComboBox denominACION 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   6495
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Razon Social"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   8160
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   11465
            MinWidth        =   11465
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "15/04/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:25 a.m."
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_recibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim EXISTE As String

Dim Fiscalrnc As Driver





Private Sub carga_comp_pendiente()
vta_recibo1.armagrid
ic = Space$(10)
id = Space$(10)
n = Space$(10)

QUERY = "select * from vta_02, vta_06 where  [estado_pago] = 'N' and [id_cliente] = " & denominACION.ItemData(denominACION.ListIndex) & " and [cta_cte] <> 'N' and vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and [contado] = 'N' and [sucursal_ingreso] = vta_06.[sucursal]"
QUERY = QUERY & " order by fecha"
Set rs = New ADODB.Recordset
rs.Open QUERY, cn1
While Not rs.EOF
   cot = rs("cotizacion_dolar")
   If rs("vta_02.moneda") = "P" Then
     sp = rs("subtotal")
     tp = rs("total")
     td = rs("total_otra_moneda")
   Else
     sp = rs("subtotal") * cot
     tp = rs("total_otra_moneda")
     td = rs("total")
   End If
   
   si2 = Format$(rs("saldo_impago02"), "#####0.00")
   
   If rs("cta_cte") = "D" Then
     RSet ic = Format$(tp, "######0.00")
     RSet n = Format$(sp, "######0.00")
     RSet id = Format$(td, "######0.00")
     si2 = Format$(rs("saldo_impago02"), "#####0.00")
   
   Else
     RSet ic = Format$(-tp, "######0.00")
     RSet n = Format$(-sp, "######0.00")
     RSet id = Format$(-td, "######0.00")
     si2 = Format$(-rs("saldo_impago02"), "#####0.00")
   
   End If
   tc = Format$(rs("letra"), "@")
   sc = Format$(rs("vta_02.sucursal"), "0000")
   nc = Format$(rs("num_comp"), "00000000")
   fc = Format$(rs("fecha"), "dd/mm/yyyy")
   cc = "(" & Format$(rs("vta_02.id_tipocomp"), "000") & ")"
   ni = Format$(rs("num_int"), "00000000")
   d = RTrim$(Format$(rs("aBREVIATURA"), "@@@@@@@@@@!"))
   tipoc = d & " " & cc & tc & " " & sc & "-" & nc
   vta_recibo1.msf1.AddItem "" & Chr$(9) & fc & Chr$(9) & tipoc & Chr$(9) & ic & Chr$(9) & ni & Chr$(9) & n & Chr$(9) & id & Chr$(9) & Format$(rs("vta_02.id_tipocomp"), "000") & Chr$(9) & si2 & Chr$(9) & ""
   rs.MoveNext
Wend
Set rs = Nothing
End Sub




Private Sub c_sucursal_LostFocus()
If c_sucursal.ListIndex < 0 Then
  c_sucursal.ListIndex = buscaindice(c_sucursal, para.punto_venta_usuario)
End If
sucursal = Format$(c_sucursal, "0000")
t_numop = ""
End Sub

Private Sub Check1_Click()
Check3 = False
End Sub

Private Sub Check3_Click()
Check1 = False
End Sub

Private Sub Command1_Click()
vta_COMPVARIOS.Show
End Sub

Private Sub Command2_Click()
vta_facturacion.Show
End Sub

Private Sub Command5_Click()
vta_clientes.t_id = denominACION.ItemData(denominACION.ListIndex)
vta_clientes.carga
vta_clientes.Show

End Sub

Private Sub Confirma_Click()
Call bloquea_comp
If denominACION.ItemData(denominACION.ListIndex) > 1 Then
If estadocaja(t_fecha) = "A" Then
 J = MsgBox("Confirma Operacion para Recibo", 4)
 If J = 6 Then
  If verificaperiodog(t_fecha) = "A" Then
   If Val(total) > 0 Then
    If Val(sucursal) = glo.sucursalf Then
      Set cl_fiscal = New fiscal
      cl_fiscal.carga (Val(sucursal))
      
      
      
      If cl_fiscal.imprimerbo = "S" Then
         'Rbo como doc fiscal
          
          Select Case cl_fiscal.idmodelo
            Case Is = 21 'tm-2000
               resulta = imprime_rbofiscal2
            
            Case Is = 22 'lx300
               resulta = imprime_rbofiscal2
               
            Case Is = 24 'tm-900 nuevo
               resulta = imprime_rbofiscal22 'nofiscal
            
            End Select
            
            If resulta Then
               If Val(t_numop) > 0 Then
                 espere.ProgressBar1.Value = 5
                 espere.Label1 = "Espere... Grabando Comprobante Fiscal"
                 Call graba
                 seguir = False
               Else
                 MsgBox ("El recibo no puede grabarse, necesita emitirlo nuevamente")
                 seguir = False
               End If
             Else
               MsgBox ("Error al Imprimir el Comprobante.")
              
             End If
             Unload espere
     Else
          'recibo como doc no fiscal
           Select Case cl_fiscal.idmodelo
            Case Is = 21 'tm-2000
               resulta = imprime_rbofiscal
            
            Case Is = 22 'lx300
               resulta = imprime_rbofiscal
               
            Case Is = 24 'tm-900 nuevo
               resulta = imprime_rbofiscal22 'nofiscal
            
            End Select
             
            
            
            
            
            If resulta Then
             If Val(t_numop) > 0 Then
               espere.ProgressBar1.Value = 5
               espere.Label1 = "Espere... Grabando Comprobante Fiscal"
               Call graba
               seguir = False
             Else
                MsgBox ("El recibo no puede grabarse, necesita emitirlo nuevamente")
                seguir = False
             End If
            Else
               MsgBox ("Error al Imprimir el Comprobante.")
             End If
             Unload espere
          
      End If
      Set cl_fiscal = Nothing
   Else
    'recibo comun
     Call graba
   
     
   End If
   
      'nota de credito
   If Val(t_diferencia) > 0 And para.ncenrecibo = "S" Then
      J = MsgBox("El Importe a pagar es mayor al importe ingresado. Desea generar Nota de Credito por descuento?", 4)
      If J = 6 Then
        Set cl_compvta = New comprobantes_venta
        cl_compvta.sucursal = Val(c_sucursal)
        cl_compvta.actual (3)
       
        
        If Val(sucursal) = glo.sucursalf Then
           t_sucnc = Format$(glo.sucursalf, "0000")
           
           
           Call fiscal2
        
        
        Else
          t_sucnc = Format$(Val(c_sucursal), "0000")
          t_subtotalnc = Format$(Val(t_diferencia) / (1 + (para.tasageneral / 100)), "#####0.00")
          t_totalnc = t_diferencia
          t_ivanc = Format$(Val(t_totalnc) - Val(t_subtotalnc), "#####0.00")
          cl_compvta.SACANUMCOMP
          t_numnc = cl_compvta.numcomp
          cl_compvta.ACTUALIZA_NUMERADOR
          
          
          
        End If
        Set cl_compvta = Nothing
        Call grabanc
        
        'imprime nc normal
        If glo.sucursalf <> Val(sucursal) Then
          J = MsgBox("Confirma Impresion de Nota Credito por Descuento", 4)
          If J = 6 Then
             Set cl_compvta = New comprobantes_venta
             cl_compvta.cargar2 (Val(t_numintnc))
             cl_compvta.imprimir
          End If
        End If

        
      End If
      
    End If
   
   
   
   
   Call pi3
 Else
  MsgBox ("El importe del Recibo debe ser mayor a cero")
 End If
Else
  MsgBox ("Periodo Cerrado. Imposible grabar comprobante")
End If
End If
Else
  MsgBox ("No se pueden realizar operaciones de caja en la fecha indicada. Caja CERRADA")
End If
Else
 MsgBox ("No se puede utilizar el cliente de contado para realizar recibos")
End If

Call libera_comp
sv = sucursal
Call INICIALIZA2(Me)
sucursal = sv
t_numop.SetFocus


End Sub


Function imprime_rbofiscal22()
Dim CUIT As String
Dim identifica As String
Dim tpago As String
Dim t As String
Dim cliz As String
Dim dirz As String
Dim locz As String
Dim de1z As String

espere.Show
espere.Refresh
espere.ProgressBar1.Min = 0
espere.ProgressBar1.Max = 6
espere.ProgressBar1.Value = 1
espere.Label1 = "Espere... Comprobando Impresora"

Set cl_fiscal = New fiscal
cl_fiscal.carga (glo.sucursalf)
caracteresmax = cl_fiscal.caracteresmax
Set cl_fiscal = Nothing

If vta_clientes.c_iva.ItemData(vta_clientes.c_iva.ListIndex) <> 3 Then
   identifica = 0 'cuit
   CUIT = RTrim$(vta_clientes.t_cuit)
 Else
   identifica = 1 'dni
   CUIT = RTrim$(vta_clientes.t_cuit)
 End If
 
tpago = "Cta.Cte. Nro: " & Format$(denominACION.ItemData(denominACION.ListIndex), "00000")
 
ef = 0
ch = ""
For i = 1 To msf2.Rows - 1
  Select Case Val(msf2.TextMatrix(i, 0))
    Case Is = 1
       ef = ef + Val(msf2.TextMatrix(i, 6))
    Case Is = 3
       ch = ch & RTrim$(msf2.TextMatrix(i, 2)) & " "
  End Select
Next i

' Call NULOS(t_remito)
 espere.ProgressBar1.Value = 2
 espere.Label1 = "Espere... Abriendo Comprobante Fiscal:" & c_tipocomp
 
 tipocompfz = 14 ' recibo
 
 'On Error GoTo errf
 cliz = textofiscal(Left$(vta_clientes.t_cli & " ", caracteresmax))
 dirz = textofiscal(Left$(vta_clientes.t_direccion & " ", caracteresmax))
 locz = textofiscal(Left$(vta_clientes.t_localidad & " ", caracteresmax))
 tivacz = vta_clientes.t_codfiscal
 
 
 'abrir recibo
 
 On Error GoTo DepuraErrores
 If Not Fiscalrnc.Inicializar Then
    Err.Raise Fiscalrnc.Error, "", Fiscalrnc.ErrorDesc
  End If
  
  Fiscalrnc.CancelarComprobante


  'datos del cliente
 If Not Fiscalrnc.DatosCliente(cliz, identifica, CUIT, tivacz, dirz) Then
      Err.Raise Fiscalrnc.Error, "", Fiscalrnc.ErrorDesc
 End If
 
 If Not Fiscalrnc.AbrirComprobante(tipocompfz) Then
     Err.Raise Fiscalrnc.Error, "", Fiscalrnc.ErrorDesc
  End If
  
 'envia items a facturar
espere.ProgressBar1.Value = 3
espere.Label1 = "Espere... Imprimiendo Pago"
 
'If Not Fiscalrnc.ImprimirTextoNoFiscal("Forma de Pago") Then
'      Err.Raise Fiscalrnc.Error, "", Fiscalrnc.ErrorDesc
' End If
   
   pf = "Pago Fact."
   For i = 1 To msf1.Rows - 1
      pf = pf & msf1.TextMatrix(i, 1) & " "
   Next i
    
    
   de1z = textofiscal(Left$(pf, caracteresmax))
          
   precio = Val(total)
          
   If Not Fiscalrnc.ImprimirItem2g("", 1, precio, 0, 0, 1, "0", 1, "", "", 0) Then
             Err.Raise Fiscalrnc.Error, "", Fiscalrnc.ErrorDesc
    End If
    
    If Not Fiscalrnc.ImprimirConceptoRecibo("A cargo " & tpago) Then
     Err.Raise Fiscalrnc.Error, "", Fiscalrnc.ErrorDesc
    End If
    
    
    't_subtotal = Fiscalrnc.subtotal.MontoNeto
    't_iva = Fiscalrnc.subtotal.MontoIVA
     T_TOTAL = Fiscalrnc.subtotal.MontoVentas
 
 espere.ProgressBar1.Value = 4
  espere.Label1 = "Espere... Cerrando Comprobante Fiscal"

  Fiscalrnc.CerrarComprobante
  
  t_numcomp = Format$(Fiscalrnc.UltimoComprobante(tipocompfz), "00000000")
  Fiscalrnc.Finalizar
  
  imprime_facturafiscal2 = True
 
    
 Exit Function
DepuraErrores:
  'Fiscalrnc.Finalizar
  MsgBox Fiscalrnc.ErrorDesc
  imprime_facturafiscal2 = False
  Exit Function
 

 
End Function
Sub bloquea_comp()
Frame1.Enabled = False
Frame3.Enabled = False
Frame4.Enabled = False
Frame5.Enabled = False
Frame6.Enabled = False
End Sub
Sub libera_comp()
Frame1.Enabled = True
Frame3.Enabled = True
Frame4.Enabled = True
Frame5.Enabled = True
Frame6.Enabled = True

End Sub
Function imprime_rbofiscal2()
Dim CUIT As String
Dim identifica As String
Dim tpago As String
Dim t As String
Dim cliz As String
Dim dirz As String
Dim locz As String
Dim de1z As String

espere.Show
espere.Refresh
espere.ProgressBar1.Min = 0
espere.ProgressBar1.Max = 6
espere.ProgressBar1.Value = 1
espere.Label1 = "Espere... Comprobando Impresora"

Set cl_fiscal = New fiscal
cl_fiscal.carga (glo.sucursalf)
caracteresmax = cl_fiscal.caracteresmax
Set cl_fiscal = Nothing

If vta_clientes.t_codfiscal <> "F" Then
   identifica = "CUIT"
   'CUIT = Mid$(t_cuit, 1, 2) & Mid$(t_cuit, 4, 8) & Mid$(t_cuit, 13, 1)
    CUIT = RTrim$(vta_clientes.t_cuit)
 Else
   identifica = "DNI"
   CUIT = "0"
 End If
 
tpago = "Cta.Cte. Nro: " & Format$(denominACION.ItemData(denominACION.ListIndex), "00000")
 
ef = 0
ch = ""
For i = 1 To msf2.Rows - 1
  Select Case Val(msf2.TextMatrix(i, 0))
    Case Is = 1
       ef = ef + Val(msf2.TextMatrix(i, 6))
    Case Is = 3
       ch = ch & RTrim$(msf2.TextMatrix(i, 2)) & " "
  End Select
Next i

' Call NULOS(t_remito)
 espere.ProgressBar1.Value = 2
 espere.Label1 = "Espere... Abriendo Comprobante Fiscal:" & c_tipocomp
 
 't_codfiscal = "I"
 cliz = textofiscal(Left$(RTrim$(denominACION) & "-", caracteresmax))
 dirz = textofiscal(Left$(RTrim$(vta_clientes.t_direccion & "-"), caracteresmax))
 locz = textofiscal(Left$(RTrim$(vta_clientes.t_localidad & "-"), caracteresmax))
 r = epson1.OpenInvoice("L", "C", "X", "1", "P", "17", "I", vta_clientes.t_codfiscal, cliz, " ", identifica, CUIT, "N", dirz, locz, tpago, "Efectivo: $" & Format$(ef, "0.00"), Left$("Ch.Nro: " & Format$(ch, "0000000000"), caracteresmax), "C")
 
 
 'r = epson1.OpenInvoice("L", "C", "X", "1", "P", "17", "I", vta_clientes.t_codfiscal, Left$(RTrim$(denominACION) & "-", caracteresmax), " ", identifica, CUIT, "N", "1-", "2-", "3.", "1r..", "2r...", "C")
 
 'envia items a facturar
espere.ProgressBar1.Value = 3
espere.Label1 = "Espere... Imprimiendo Pago"
 
 If r Then
   pf = "Pago Fact."
   For i = 1 To msf1.Rows - 1
      pf = pf & msf1.TextMatrix(i, 1) & " "
   Next i
    r = epson1.SendInvoiceItem(Left$(pf, caracteresmax), "0", Format$(Val(total) * 100, "000000000"), "0000", "M", "0", "0", " ", Left$(Detalle & " ", caracteresmax), " ", "0", "0")
 Else
       Call verificaerrfiscal(epson1.FiscalStatus, epson1.PrinterStatus)
 End If
 
 'no realiza pago copmprobante y automaticanmene cuando cierro se genera un pago pr el total
  
 'subtotal para obtener el importe neto, iva y total impreso en la factura
espere.ProgressBar1.Value = 4
espere.Label1 = "Espere... Cerrando Comprobante Fiscal"

 
 
If r Then
   r = epson1.CloseInvoice("L", "X", " ")
Else
  Call verificaerrfiscal(epson1.FiscalStatus, epson1.PrinterStatus)
 End If

If r Then
   t_numop = epson1.AnswerField_3
Else
  Call verificaerrfiscal(epson1.FiscalStatus, epson1.PrinterStatus)
 End If
imprime_rbofiscal2 = r
    
   'si hay error tratarlo en un proceso global de errores fiscales

End Function

Function imprime_rbofiscal() As Boolean
Dim CUIT As String
Dim identifica As String
Dim tpago As String

Set cl_cli = New Clientes
cl_cli.carga (denominACION.ItemData(denominACION.ListIndex))
tpago = asteriscos(Format$(Val(total), "######0.00"), 15)
espere.Show
espere.Refresh
espere.ProgressBar1.Min = 0
espere.ProgressBar1.Max = 5
espere.ProgressBar1.Value = 1
espere.Label1 = "Espere... Abriendo Comprobante Fiscal"

'variables copn renglones
nc = "Nro: " & Format$(sucursal, "0000") & "-" & Format$(t_numop, "00000000")
cli = "(" & Format$(cl_cli.id, "00000") & ") " & Left$(cl_cli.razonsocial, 80)
lloc = Format$(Left$(cl_cli.localidad, 40), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!") & "   Cuit:" & cl_cli.abreviatura_tipoiva & " " & cl_cli.CUIT

r = epson1.OpenNoFiscal
espere.ProgressBar1.Value = 2
espere.Label1 = "Espere... Imprimiendo Recibo"

If r Then
   r = epson1.SendNoFiscalText("RECIBO OFICIAL (NO VALIDO COMO FACTURA)  " & nc)
Else
   MsgBox ("Error F001 al Inicializar Comprobante. Estado Fiscal: " & epson1.FiscalStatus & "  Estado Impresor. " & epson1.PrinterStatus)
End If
'If r Then r = epson1.SendNoFiscalText(" ")
If r Then r = epson1.SendNoFiscalText("Recibimos de: " & cli)
If r Then r = epson1.SendNoFiscalText("Domicilio...: " & cl_cli.direccion)
If r Then r = epson1.SendNoFiscalText("Localidad...: " & lloc)
If r Then r = epson1.SendNoFiscalText(" ")
If r Then r = epson1.SendNoFiscalText("--------------------------------------------------------------------------------")
If r Then r = epson1.SendNoFiscalText("La suma de $: " & tpago)
If r Then r = epson1.SendNoFiscalText("--------------------------------------------------------------------------------")
If r Then r = epson1.SendNoFiscalText(" ")
'If r Then r = epson1.SendNoFiscalText("Forma Pago")
If r Then r = epson1.SendNoFiscalText("-------------------------------------------------")
If r Then r = epson1.SendNoFiscalText("Forma Pago   /   Importe ")
If r Then r = epson1.SendNoFiscalText("-------------------------------------------------")
For i = 1 To msf2.Rows - 1
  If r Then
   If cl_fiscal.caracteresmax < 39 Then
     r = epson1.SendNoFiscalText(Left$(msf2.TextMatrix(i, 1), 10) & " " & Left$(msf2.TextMatrix(i, 3), 18) & " " & Format$(Val(msf2.TextMatrix(i, 6)), "#####0.00"))
   Else
     r = epson1.SendNoFiscalText(Left$(msf2.TextMatrix(i, 1), 10) & " " & Left$(msf2.TextMatrix(i, 3), 18) & " Nro. " & Left$(msf2.TextMatrix(i, 2), 10) & " $" & Format$(Val(msf2.TextMatrix(i, 6)), "#####0.00"))
   End If
  Else
     i = msf2.Rows
  End If
Next i
If r Then r = epson1.SendNoFiscalText(" ")
If r Then r = epson1.SendNoFiscalText("-------------------------------------------------")
If r Then r = epson1.SendNoFiscalText("Comprobantes Aplicados  ")
If r Then r = epson1.SendNoFiscalText("-------------------------------------------------")
For i = 1 To msf1.Rows - 1
  If r Then
     r = epson1.SendNoFiscalText(msf1.TextMatrix(i, 0) & "  " & msf1.TextMatrix(i, 1))
  Else
     i = msf1.Rows
  End If
Next i
If r Then r = epson1.SendNoFiscalText(" ")
If r Then r = epson1.SendNoFiscalText("Observaciones..: " & t_obs)
If r Then r = epson1.SendNoFiscalText("                                            Recibi Conforme:___________________ ")

espere.ProgressBar1.Value = 3
espere.Label1 = "Espere... Cerrando Comprobante Fiscal"

If r Then r = epson1.CloseNoFiscal
   
imprime_rbofiscal = r
    
Set cl_cli = Nothing
'si hay error tratarlo en un proceso global de errores fiscales
   
End Function

Private Sub denominACION_LostFocus()
espere.Show
espere.Label1 = "Inicializando Comprobante"
If denominACION.ListIndex < 0 Then
  If Val(denominACION) > 0 Then
    denominACION.ListIndex = buscaindice(denominACION, Val(denominACION))
  Else
    denominACION.ListIndex = 0
  End If
  
End If
vta_clientes.t_id = denominACION.ItemData(denominACION.ListIndex)
vta_clientes.carga

'calcula saldo
Set cl_cli = New Clientes
cl_cli.carga (denominACION.ItemData(denominACION.ListIndex))
If cl_cli.id > 1 Then
    t_saldo21 = Format$(cl_cli.saldo(True, Now, True), "######0.00")
Else
    t_saldo21 = 0
End If
Set Clientes = Nothing



Unload espere
End Sub
Private Sub detalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Val(total) <> Val(t_ingresado) Then
     MsgBox ("El importe Ingresado no coincide con el total Aplicado")
   End If
     confirma.Enabled = True
     confirma.SetFocus
   
End If



End Sub


Private Sub fdolar_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)
End Sub

Private Sub fdolar_LostFocus()
If Val(fdolar) <= 1 Then
   fdolar = "1.00"
Else
   fdolar = Format$(fdolar, "####0.000")
End If
t_totald = Format$(Val(total) / Val(fdolar), "#####0.00")

End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 9
msf1.ColWidth(0) = 1200
msf1.ColWidth(1) = 2500
msf1.ColWidth(2) = 1100
msf1.ColWidth(3) = 1100
msf1.ColWidth(4) = 1100
msf1.ColWidth(5) = 1100
msf1.ColWidth(6) = 500
msf1.ColWidth(7) = 1100
msf1.ColWidth(8) = 1100
msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = "Comprobante"
msf1.TextMatrix(0, 2) = "Total $"
msf1.TextMatrix(0, 3) = "Num.Int."
msf1.TextMatrix(0, 4) = "Neto $"
msf1.TextMatrix(0, 5) = "Total U$s"
msf1.TextMatrix(0, 6) = "Tipo"
msf1.TextMatrix(0, 7) = "Saldo $"
msf1.TextMatrix(0, 8) = "Aplicar $"

End Sub

Sub armagrid2()
msf2.clear
msf2.Rows = 1
msf2.Cols = 10
msf2.ColWidth(0) = 600
msf2.ColWidth(1) = 1200
msf2.ColWidth(2) = 1200
msf2.ColWidth(3) = 2500
msf2.ColWidth(4) = 1700
msf2.ColWidth(5) = 1700
msf2.ColWidth(6) = 1000
msf2.ColWidth(7) = 1000
msf2.ColWidth(8) = 1000
msf2.ColWidth(9) = 1000

msf2.TextMatrix(0, 0) = "Cod."
msf2.TextMatrix(0, 1) = "Forma Pago"
msf2.TextMatrix(0, 2) = "Num.Cheque"
msf2.TextMatrix(0, 3) = "Detalle/Banco"
msf2.TextMatrix(0, 4) = "Sucursal"
msf2.TextMatrix(0, 5) = "Titular"
msf2.TextMatrix(0, 6) = "Importe"
msf2.TextMatrix(0, 7) = "Fecha Dif."
msf2.TextMatrix(0, 8) = "Num.Int."
msf2.TextMatrix(0, 9) = "Cuenta"


End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 11)
  
End Select
End Sub

Sub graba()
' On Error GoTo ERRORGRABA
numint = saca_ultnumero_int_comp("V")
    
Set cl_compvta = New comprobantes_venta
cl_compvta.sucursal = Val(sucursal)
cl_compvta.actual (50)
cl_compvta.letra = "R"
cl_compvta.numcomp = Val(t_numop)
cl_compvta.sucursal = Val(sucursal)
cl_compvta.ACTUALIZA_NUMERADOR
      
Set cl_cli = New Clientes
cl_cli.carga (denominACION.ItemData(denominACION.ListIndex))
 t2d = Val(t_totald)
 t2p = Val(total)
 If Check1 = 1 Then
   t2p = 0
 End If
 
 If Check3 = 1 Then
   t2d = 0
 End If
        
 cn1.BeginTrans
        
QUERY = "INSERT INTO vta_02([num_int], [sucursal], [num_comp], [letra], [id_tipocomp], [id_cliente], [fecha], [id_usuario], [subtotal], [impuestos], [iva], [total], [estado], [id_cuenta], [stock], [cta_cte], " & _
"[grabado], [estado_pago], [recibo_Pago], [observaciones], [cotizacion_dolar], [total_otra_moneda], [moneda], [id_vendedor], [VENTA], [CONTADO], [fecha_vto], [id_actividad], [alicuota_ib], " & _
"[alicuota_perc_iva], [canje_cereal], [total_bultos], [valor_declarado], [transporte], [direccion_transp], [cuit_transp], [perc_ss], [sucursal_ingreso], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02], [chofer02], [dominio02], [dominio_acoplado02], " & _
"[cae], [cae_vence], [tipo_op])"

QUERY = QUERY & " VALUES (" & numint & ", " & Val(sucursal) & ", " & Val(t_numop) & ", 'R', 50, " & denominACION.ItemData(denominACION.ListIndex) & ", '" & t_fecha & "', " & para.id_usuario & _
", 0, 0, 0, " & t2p & ", 'A', " & para.cuenta_ventas & ", '" & cl_compvta.STOCK & "', '" & cl_compvta.ctacte & "', '" & cl_compvta.grabado & "', 'S', '0000-00000000', '" & Detalle & " ', " & Val(fdolar) & _
", " & t2d & ", 'P', " & cl_cli.idvendedor & ", '" & cl_compvta.venta & "', 'N', '" & t_fecha & "' ,1 , 0, 0, 0, 0, 0, ' ', ' ', ' ', 0, " & Val(c_sucursal) & ", '" & Left$(cl_cli.razonsocial, 50) & "', '" & _
Left$(cl_cli.direccion, 50) & "', '" & Left$(cl_cli.CUIT, 20) & "', '" & Left$(cl_cli.localidad, 50) & "', " & cl_cli.idtipoiva & ", ' ', ' ', ' ', 'u2', '01/01/2018', 2)"

cn1.Execute QUERY
      
      For i = 1 To msf2.Rows - 1
         If Val(msf2.TextMatrix(i, 0)) = 3 Then
                'ch. terceros
                q = "select * from cyb_03"
                Set rs = New ADODB.Recordset
                rs.Open q, cn1, adOpenDynamic, adLockOptimistic
                rs.AddNew
                 rs("fecha_emision") = t_fecha
                 rs("num_cheque") = Val(msf2.TextMatrix(i, 2))
                 rs("banco") = msf2.TextMatrix(i, 3)
                 rs("sucursal") = msf2.TextMatrix(i, 4)
                 rs("titular") = msf2.TextMatrix(i, 5)
                 rs("importe") = Val(msf2.TextMatrix(i, 6))
                 rs("estado") = "C"
                 rs("fecha_dif") = msf2.TextMatrix(i, 7)
                 rs("origen") = Left$(denominACION, 50)
                 rs("destino") = " "
                 rs("num_mov_banco_i") = 0
                 rs("num_mov_banco_e") = 0
                 rs("num_int_op") = 0
                 rs("num_int_rbo") = numint
                 rs("fecha_salida") = t_fecha
                 rs("fecha_ingreso") = t_fecha
                 rs("tipo_salida") = "C"
                rs.Update
                
                qr = "SELECT @@IDENTITY AS NewID"
                Set rs = cn1.Execute(qr)
                numintch = rs.Fields("NewID").Value

                
                Set rs = Nothing
         
         Else
           numintch = 0
         End If
         
         
         If Val(msf2.TextMatrix(i, 0)) = 4 Then
                q = "select * from cyb_04"
                Set rs = New ADODB.Recordset
                rs.Open q, cn1, adOpenDynamic, adLockOptimistic
                rs.AddNew
                 rs("id_banco") = Val(msf2.TextMatrix(i, 8))
                 rs("fecha") = msf2.TextMatrix(i, 7)
                 rs("importe") = Val(msf2.TextMatrix(i, 6))
                 rs("id_tipomov") = 60 'transf
                 rs("fecha_dif") = msf2.TextMatrix(i, 7)
                 rs("ubicacion") = "H"
                 rs("entro") = "N"
                 rs("fecha_acreed") = msf2.TextMatrix(i, 7)
                 rs("num_comp") = Val(msf2.TextMatrix(i, 2))
                 rs("detalle") = "Transf." & Left$(msf2.TextMatrix(i, 5), 30)
                 rs("modulo") = "V"
                 rs("num_mov_int") = numint
                 rs("id_tipodbcr") = 1
                rs.Update
                
                Set rs = Nothing
         End If
         
         
         q = "select * from cyb_01 where [id_forma_pago] = " & Val(msf2.TextMatrix(i, 0))
         Set rs = New ADODB.Recordset
         rs.Open q, cn1
         If Not rs.EOF And Not rs.BOF Then
          If rs("CAJA") = "S" Then
             ctach = rs("id_cuenta_cont")
             QUERY = "INSERT INTO cyb_05([id_cuenta_caja], [id_cuenta_contra], [descripcion], [importe], [ubicacion], [fecha], [num_mov_int], [modulo], [operacion], [id_forma_pago], [num_int_ch_terc], [id_usuario])"
             QUERY = QUERY & " VALUES (" & ctach & ", " & para.cuenta_deudores & ", '" & RTrim$(Left$(denominACION, 49)) & " ', " & Val(msf2.TextMatrix(i, 6)) & ", 'D', '" & t_fecha & "', " & numint & ", 'V', 'Rbo." & Format$(sucursal, "0000") & "-" & Format$(t_numop, "00000000") & "', " & Val(msf2.TextMatrix(i, 0)) & ", " & numintch & ", " & para.id_usuario & ")"
             cn1.Execute QUERY
          End If
         End If
         Set rs = Nothing

                 
        'formas de pago
        QUERY = "INSERT INTO vta_04([num_int], [secuencia], [id_formapago], [formapago], [num_ch], [detalle_banco], [sucursal], [titular], [importe], [fecha_dif], [num_int_fp])"
        QUERY = QUERY & " VALUES (" & numint & ", " & i & ", " & Val(msf2.TextMatrix(i, 0)) & ", '" & Left$(RTrim$(msf2.TextMatrix(i, 1)), 9) & " ', " & Val(msf2.TextMatrix(i, 2)) & ", '" & RTrim$(msf2.TextMatrix(i, 3)) & " ', '" & RTrim$(msf2.TextMatrix(i, 4)) & " ', '" & RTrim$(msf2.TextMatrix(i, 5)) & " ', " & Val(msf2.TextMatrix(i, 6)) & ", '" & RTrim$(msf2.TextMatrix(i, 7)) & " ', " & numintch & ")"
        cn1.Execute QUERY

      Next i
     
      
      'actualiza comprobantes aplicados
      
      For i = 1 To msf1.Rows - 1
        Set rs = New ADODB.Recordset
        q = "select * from vta_02 where [num_int] = " & Val(msf1.TextMatrix(i, 3))
        rs.Open q, cn1, adOpenDynamic, adLockOptimistic
        If Not rs.BOF And Not rs.EOF Then
          If Val(msf1.TextMatrix(i, 7)) > 0 Then
             rs("estado_pago") = "N"
          Else
             rs("estado_pago") = "P"
          End If
          
          rs("saldo_impago02") = Val(msf1.TextMatrix(i, 7))
          rs("recibo_pago") = Format$(sucursal, "0000") & "-" & Format$(t_numop, "00000000")
          rs.Update
          
          QUERY = "INSERT INTO vta_010([num_int_comp], [num_int_rbo], [importe_pagado], [saldo_comprobante])"
          QUERY = QUERY & " VALUES (" & Val(msf1.TextMatrix(i, 3)) & ", " & numint & ", " & Val(msf1.TextMatrix(i, 8)) & ", " & Val(msf1.TextMatrix(i, 7)) & ")"
          cn1.Execute QUERY
     
          
        End If
        Set rs = Nothing
      Next i
        
      
   'contabilidad
   If Generaasientosauto Then
      
      If cl_compvta.contabilidad <> "N" Then
         numintcgr = saca_ultnumero_int_comp("G")
         cta = para.cuenta_deudores
         u1 = cl_compvta.contabilidad
         If u1 = "D" Then
           u2 = "H"
         Else
           u2 = "D"
         End If
         
         Set rs = New ADODB.Recordset
         q = "select * from c_01 where [id_cuenta] = " & cta
         rs.Open q, cn1
         If Not rs.EOF And Not rs.BOF Then
           dcta = rs("descripcion")
         Else
           dcta = "Cuenta Inexistente"
         End If
         Set rs = Nothing
         
         'grabo asiento
         QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Cobranza] " & cl_compvta.abreviatura & " " & t_letra & Format$(Val(sucursal), "0000") & "-" & Format$(Val(t_numop), "00000000") & "', 'V', " & numint & ", " & Val(total) & ", " & Val(total) & ", " & para.id_usuario & ", '" & Left$(RTrim$(denominACION), 50) & "')"
         cn1.Execute QUERY
      
         'cuenta madre deudores
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", 1, " & cta & ", '" & u1 & "', " & Val(total) & ", 'Rbo. Nro." & Format$(Val(sucursal), "0000") & "-" & Format$(Val(t_numop), "00000000") & "')"
         cn1.Execute QUERY
      
         'formas de pago
         ic = 2
         For i = 1 To msf2.Rows - 1
            
              d = Left$(RTrim$(msf2.TextMatrix(i, 3)), 35) & " " & msf2.TextMatrix(i, 2)
              cta = Val(msf2.TextMatrix(i, 9))
              QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
              QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cta & ", '" & u2 & "', " & Val(msf2.TextMatrix(i, 6)) & ", '" & d & "')"
              cn1.Execute QUERY
              ic = ic + 1
          
         Next i
      
      End If
     End If
      
      
      
      
      
      cn1.CommitTrans
      Set rs = Nothing
      Set cl_cli = Nothing
      
      
      'impresion de recibo
       
      If Val(sucursal) <> glo.sucursalf Then
          'If glo.sucursalf = 0 Then
             J = MsgBox("Imprime Comprobante", 4)
             If J = 6 Then
                 cl_compvta.cargar2 (numint)
                 If cl_compvta.numint > 0 Then
                   cl_compvta.imprimir
                 End If
             End If
          'Else
          '   MsgBox ("Por disposicion del AFIP teniendo una impresora fiscal definida no se permite imprimir otro tipo de comprobantes. Gracias")
          'End If
      
             
      
      Else
             J = MsgBox("Imprime Minuta para el Recibo Fiscal(Utiliza impresora Común)", 4)
             If J = 6 Then
                 cl_compvta.cargar2 (numint)
                 If cl_compvta.numint > 0 Then
                   cl_compvta.imprimeminutafiscal
                 End If
             End If
             
     
      End If
      
      
      
      Set cl_cli = Nothing
      
      
      cl_compvta.web (numint)
      
      Set cl_compvta = Nothing
   
      
   
      
Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  

End Sub


Private Sub Form_Load()
Call INICIALIZA2(Me)
sucursal = Format$(para.punto_venta_usuario, "0000")
denominACION.clear
Call carga_clientes(denominACION)
denominACION.RemoveItem 0
denominACION.ListIndex = 0
Call carga_SUCURSALES(c_sucursal)
c_sucursal.ListIndex = buscaindice(c_sucursal, para.punto_venta_usuario)

Load vta_recibo1
Load vta_recibo2
Load vta_recibo3
Load vta_recibo4
fdolar = para.cotizacion
Load vta_clientes


Set cl_fiscal = New fiscal
cl_fiscal.carga (glo.sucursalf)
If cl_fiscal.idmodelo <> 24 Then
  epson1.PortNumber = cl_fiscal.puerto
Else
 
   Set Fiscalrnc = New Driver
  Fiscalrnc.Modelo = cMODELO
  Fiscalrnc.puerto = cPUERTO
  Fiscalrnc.baudios = cBAUDIOS
End If
Set cl_fiscal = Nothing


 



End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload vta_recibo1
Unload vta_recibo2
Unload vta_recibo3
Unload vta_recibo4
Unload vta_clientes
End Sub





Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[INS] Agrega - [ENTER] Continua - [F5] Elimina - "
If msf1.Rows > 0 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If


End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
 If msf1.Rows > 2 Then
    msf1.RemoveItem (msf1.Row)
 Else
   Call armagrid
 End If
End If



If KeyCode = vbKeyInsert Then
   espere.Show
   espere.Label1 = "Cargando comprobantes impagos"
   espere.Refresh
   Call carga_comp_pendiente
   Unload espere

   vta_recibo1.Show
End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  t_pago = Format$(suma_msflexgrid(Me.msf1, 2), "######0.00")
  total = Format$(Val(t_pago) - Val(retencion))
  If msf1.Rows > 1 Then
    msf2.SetFocus
  Else
    total = ""
    total.SetFocus
    
  End If
End If
End Sub

Private Sub msf1_LostFocus()
  Call totales2
  msf1.FocusRect = flexFocusLight
End Sub

Private Sub msf2_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F1] Ch.Terc.  - [F2] TRansferncias - [F3] Otras formas pago - [F5] Borra - [ENTER] Continua  "
If msf2.Rows > 0 Then
  msf2.FocusRect = flexFocusNone
Else
  msf2.FocusRect = flexFocusLight
End If
t_ingresado = suma_msflexgrid(msf2, 6)
If Val(total) >= 0 Then
  t_diferencia = Format$(Val(t_pago) - Val(t_ingresado), "######0.00")
Else
  t_diferencia = Format$(Val(t_pago) + Val(t_ingresado), "######0.00")
End If
total = Format$(Val(t_ingresado), "######0.00")

End Sub

Private Sub msf2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
  vta_recibo3.Show
  vta_recibo3.t_modulo = "R"
End If

If KeyCode = vbKeyF2 Then
  vta_recibo4.Show
  vta_recibo4.t_modulo = "R"
End If


If KeyCode = vbKeyF1 Then
  vta_recibo2.Show
  vta_recibo2.t_modulo = "R"
End If


If KeyCode = vbKeyF9 Then
  If Val(total) <> Val(t_ingresado) Then
     MsgBox ("El total ingresado no coincide con el total del Comprobante")
  Else
     total.SetFocus
  End If
  
End If


If KeyCode = vbKeyF5 Then
 If msf2.Rows > 2 Then
    msf2.RemoveItem (msf2.Row)
 Else
   Call armagrid2
 End If
End If

End Sub

Private Sub msf2_LostFocus()
Call barra(Me)
t_ingresado = suma_msflexgrid(msf2, 6)
msf2.FocusRect = flexFocusLight
End Sub



Private Sub pi3()
   
   t_fecha = Format$(Now, "dd/mm/yyyy")
   fdolar = Format$(para.cotizacion, "#####0.00")
   Call armagrid
   Call armagrid2
   Set cl_compvta = New comprobantes_venta
   cl_compvta.idtipocomp = 50
   cl_compvta.sucursal = Val(sucursal)
   cl_compvta.SACANUMCOMP
   t_numop = Format$(cl_compvta.numcomp, "00000000")
   Set cl_compvta = Nothing
   confirma.Enabled = False
End Sub

Private Sub Option3_Click()
Label5(1) = "Total $"
End Sub

Private Sub Option4_Click()
Label5(1) = "Total U$s"
End Sub

Private Sub Salir_Click()
   
   Unload Me

End Sub

Private Sub sucursal_GotFocus()
 Call pi3
End Sub

Private Sub sucursal_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Unload Me
Else
  Call solonum(KeyAscii, 0)
End If


End Sub

Private Sub sucursal_LostFocus()
Set rs = New ADODB.Recordset
q = "select * from vta_06 where [sucursal] = " & Val(sucursal)
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
  Call pi3
Else
  MsgBox ("Sucursal no Habilitada")
  t_sucursal = Format$(glo.sucursal, "0000")
  Call pi3
End If
End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
      t_fecha = Format$(Now, "dd/mm/yyyy")
  End If
Else
  t_fecha = Format$(Now, "dd/mm/yyyy")
End If
t_fecha = Format$(t_fecha, "dd/mm/yyyy")
Call verifica_fechacorte(t_fecha)

End Sub

Private Sub t_numop_GotFocus()
 Call pi3
End Sub

Private Sub t_numop_LostFocus()
      q = "select * from vta_02 where [sucursal] = " & Val(sucursal) & " and [num_comp] = " & Val(t_numop) & " and [id_tipocomp] = 50"
      Set rs = New ADODB.Recordset
      'MsgBox (q)
      rs.Open q, cn1
      If rs.BOF And rs.EOF Then
         EXISTE = "N"
      Else
         MsgBox ("Recibo Existente")
         EXISTE = "S"
         sucursal.SetFocus
      End If
      
  
End Sub

Private Sub t_pago_KeyPress(KeyAscii As Integer)
  Call solonum(KeyAscii, 1)
End Sub


Private Sub t_pago_LostFocus()
Call totales
End Sub

Private Sub t_retenciones_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)
End Sub

Private Sub t_retenciones_LostFocus()
Call totales

End Sub

Private Sub t_totald_LostFocus()
   Call fm(t_totald)
   
End Sub

Private Sub totales()
      If Val(fdolar) < 1 Then
        fdolar = "1.00"
      End If
      t_pago = Format$(Val(t_pago), "#####0.00")
      'If Val(t_ingresdado) > Val(t_retenciones) Then
         total = Format$(Val(t_ingresado), "#####0.00")
         t_totald = Format$(Val(total) / Val(fdolar), "#####0.00")
      'Else
      '   total = "0.00"
      '   totald = "0.00"
      'End If

      
End Sub

Sub totales2()
  J = 1
  t_pago = 0
  t_totald = 0
  t_retenciones = 0
  T_RETD = 0
  If msf1.Rows > 1 Then
   While J <= msf1.Rows - 1
    If Val(msf1.TextMatrix(J, 6)) < 35 Then
     t_pago = Val(t_pago) + Val(msf1.TextMatrix(J, 8))
     't_totald = Val(t_totald) + Val(msf1.TextMatrix(J, 5))
    Else
      If Val(msf1.TextMatrix(J, 6)) >= 100 And Val(msf1.TextMatrix(J, 6)) <= 110 Then
         t_retenciones = Val(t_retenciones) + Val(msf1.TextMatrix(J, 2))
         'T_RETD = Val(T_RETD) + Val(msf1.TextMatrix(J, 5))
      End If
    End If
    J = J + 1
   Wend
  End If
  t_pago = Format$(Val(t_pago), "######0.00")
  t_retenciones = Format$(-Val(t_retenciones), "#####0.00")
  If Val(t_pago) > Val(t_retenciones) Then
   total = Format$(Val(t_pago) - Val(t_retenciones), "######0.00")
   t_totald = Format$(Val(total) / Val(fdolar), "######0.00")
  Else
   total = "0.00"
   totald = "0.00"
  End If
  
  
  'If Val(t_totald) >= 1 Then
  '  fdolar = Format$(Val(t_pago) / Val(t_totald), "###0.000")
  'Else
  '  fdolar = "1.000"
  'End If
End Sub
Private Sub total_LostFocus()
Call fm(total)
t_totald = Format$(Val(total) / Val(fdolar), "#####0.00")

End Sub


Sub fiscal2()
    'formato fiscal
            
Set cl_fiscal = New fiscal
cl_fiscal.carga (Val(sucursal))
      
            seguir = True
            While seguir
              If cl_fiscal.idmodelo <> 24 Then 'tm-900
                resulta = imprime_NCf2
              Else
                resulta = imprime_NCf22 'nuevo protocolo
              End If
              If resulta Then
                seguir = False
              Else
                J = MsgBox("Error al Imprimir el Comprobante. Verifique la Impresora para continuar.  Reintente o Cancele", 5)
                If J = 4 Then
                  seguir = True
                Else
                  seguir = False
                End If
              End If
              Unload espere
            Wend
Set cl_fiscal = Nothing


End Sub
Function imprime_NCf22() As Boolean
  'nuevo protocolo
  
  Dim a(5) As String

Dim CUIT As String
Dim identifica As String
Dim tpago As String
Dim t  As String
Dim de1 As String
Dim tipocompfz As String
Dim tv2 As String
Dim td As String
Dim cliz As String
Dim dirz As String
Dim locz As String
Dim de1z As String
Dim tivacz As String
Dim letraz As String
Dim rk As Boolean
Dim remitosz As String
Dim remitosz2 As String
'Dim r As Boolean
Set cl_fiscal = New fiscal
cl_fiscal.carga (glo.sucursalf)
para.z_actual = cl_fiscal.ultimo_z + 1
   If cl_fiscal.imprimenc = "S" Then
       If t_letra = "A" Then
           tipocompfz = 7
       Else
           tipocompfz = 8
       End If
    Else
        MsgBox ("La impresora fiscal no puede imprimir NC")
        imprime_NCf22 = False
         Exit Function
    End If

caracteresmax = cl_fiscal.caracteresmax
Set cl_fiscal = Nothing




espere.Show
espere.Refresh
espere.ProgressBar1.Min = 0
espere.ProgressBar1.Max = 6
espere.ProgressBar1.Value = 1
espere.Label1 = "Espere... Comprobando Impresora"
'abrir factura
If vta_clientes.c_iva.ItemData(vta_clientes.c_iva.ListIndex) <> 3 Then
   identifica = 0 'cuit
   'CUIT = Mid$(vta_clientes.t_cuit, 1, 11) '& Mid$(vta_clientes.t_cuit, 4, 8) & Mid$(vta_clientes.t_cuit, 13, 1)
    CUIT = RTrim$(vta_clientes.t_cuit)
 Else
   identifica = 1 'dni
   CUIT = RTrim$(vta_clientes.t_cuit)
 End If
 
 If Option1 = True Then
    tpago = "Cta.Cte. Nro. " & Format$(vta_clientes.t_id, "00000")
 Else
    tpago = "CONTADO"
 End If

   tv2 = " "

 espere.ProgressBar1.Value = 2
 espere.Label1 = "Espere... Abriendo Comprobante Fiscal:"
 
remitosz = " "
remitosz2 = " "



      
 
 'On Error GoTo errf
 cliz = textofiscal(Left$(vta_clientes.t_cli & " ", caracteresmax))
 dirz = textofiscal(Left$(vta_clientes.t_direccion & " ", caracteresmax))
 locz = textofiscal(Left$(vta_clientes.t_localidad & " ", caracteresmax))
 letraz = vta_clientes.t_letrafact
 tivacz = vta_clientes.t_codfiscal2
 
 
 'abrir factura
 
 On Error GoTo DepuraErrores
 If Not Fiscalrnc.Inicializar Then
    Err.Raise Fiscalrnc.Error, "", Fiscalrnc.ErrorDesc
  End If
  
  Fiscalrnc.CancelarComprobante
    
  
  
 'datos del cliente
 If Not Fiscalrnc.DatosCliente(cliz, identifica, CUIT, tivacz, dirz) Then
      Err.Raise Fiscalrnc.Error, "", Fiscalrnc.ErrorDesc
 End If
     
 
  
  
  
  If Not Fiscalrnc.AbrirComprobante(tipocompfz) Then
     Err.Raise Fiscalrnc.Error, "", Fiscalrnc.ErrorDesc
  End If
  
  
   
'envia items a facturar
espere.ProgressBar1.Value = 3
espere.Label1 = "Espere... Imprimiendo Productos"
 

 If letraz = "A" Then
      pu = Val(t_diferencia) / (1 + (para.tasageneral / 100))
  Else
      pu = Val(t_diferencia)
  End If
  
   If Not Fiscalrnc.ImprimirItem2g("Descuento", 1, pu, para.tasageneral, 0, IFUniversal.Gravado, 0, 1, "", "", 0) Then
             Err.Raise Fiscalrnc.Error, "", Fiscalrnc.ErrorDesc
    End If
  
 'pagos
  espere.Label1 = "Espere... Grabando Pagos"
  
  
  
 
 
      t_subtotalnc = Fiscalrnc.subtotal.MontoNeto
      t_ivanc = Fiscalrnc.subtotal.MontoIVA
      t_totalnc = Fiscalrnc.subtotal.MontoVentas
      

 
 td = "Cta. Cte. Nro. " & Format$(denominACION.ItemData(denominACION.ListIndex), "00000")
mp = Val(T_TOTAL)
dp = "T"
If Not Fiscalrnc.ImprimirPago2g(td, Format$(mp, "######0.00"), "", IFUniversal.CuentaCorriente, 1, "", "") Then
       Err.Raise Fiscalrnc.Error, "", Fiscalrnc.ErrorDesc
End If


  
 'subtotal para obtener el importe neto, iva y total impreso en la factura
espere.ProgressBar1.Value = 4
espere.Label1 = "Espere... Cerrando Comprobante Fiscal"

  espere.Label1.Refresh
  Fiscalrnc.CerrarComprobante
  
 
  t_numnc = Format$(Fiscalrnc.UltimoComprobante(tipocompfz), "00000000")
  Fiscalrnc.Finalizar
  
  imprime_NCf22 = True
 
    
 Exit Function
DepuraErrores:
  'Fiscalrnc.Finalizar
  MsgBox Fiscalrnc.ErrorDesc
  imprime_NCf22 = False
  Exit Function
   
End Function


Function imprime_NCf2() As Boolean
Dim a(5) As String

Dim CUIT As String
Dim identifica As String
Dim tpago As String
Dim t  As String
Dim de1 As String
Dim tipocompfz As String
Dim tv2 As String
Dim td As String
Dim cliz As String
Dim dirz As String
Dim locz As String
Dim de1z As String
Dim tivacz As String
Dim letraz As String
Dim rk As Boolean
Dim remitosz As String
Dim remitosz2 As String
'Dim r As Boolean

'parametros fiscales
Set cl_fiscal = New fiscal
cl_fiscal.carga (glo.sucursalf)
para.z_actual = cl_fiscal.ultimo_z + 1
    
    If cl_fiscal.imprimenc = "S" Then
        tipocompfz = cl_fiscal.CODNC
    Else
        MsgBox ("La impresora fiscal no puede imprimir NC")
        imprime_facturafiscal = False
         Exit Function
    End If
caracteresmax = cl_fiscal.caracteresmax
Set cl_fiscal = Nothing




espere.Show
espere.Refresh
espere.ProgressBar1.Min = 0
espere.ProgressBar1.Max = 6
espere.ProgressBar1.Value = 1
espere.Label1 = "Espere... Comprobando Impresora"
'abrir factura
If vta_clientes.c_iva.ItemData(vta_clientes.c_iva.ListIndex) <> 3 Then
   identifica = "CUIT"
   CUIT = RTrim$(vta_clientes.t_cuit)
 Else
   identifica = "DNI"
   CUIT = RTrim$(vta_clientes.t_cuit)
 End If
 
 tpago = "Cta.Cte. Nro. " & Format$(vta_clientes.t_id, "00000")
 
 espere.ProgressBar1.Value = 2
 espere.Label1 = "Espere... Abriendo Comprobante Fiscal:" & c_tipocomp
 
remitosz = " "
remitosz2 = " "

      
 
 'On Error GoTo errf
 cliz = textofiscal(Left$(vta_clientes.t_cli & " ", caracteresmax))
 dirz = textofiscal(Left$(vta_clientes.t_direccion & " ", caracteresmax))
 locz = textofiscal(Left$(vta_clientes.t_localidad & " ", caracteresmax))
 letraz = vta_clientes.t_letrafact
 tivacz = vta_clientes.t_codfiscal
 rk = epson1.OpenInvoice(tipocompfz, "C", letraz, "1", "P", "17", "I", tivacz, cliz, " ", identifica, CUIT, "N", dirz, locz, tpago, "0", " ", "C")
 If rk Then

Else
   Call verificaerrfiscal(epson1.FiscalStatus, epson1.PrinterStatus)
 
End If
 
 'envia items a facturar
espere.ProgressBar1.Value = 3
espere.Label1 = "Espere... Imprimiendo Productos"
 If letraz = "A" Then
      pu = Val(t_diferencia) / (1 + (para.tasageneral / 100))
  Else
      pu = Val(t_diferencia)
  End If
  rk = epson1.SendInvoiceItem("Descuentos Otorgados", Format$(1 * 1000, "00000000"), Format$(pu * 100, "000000000"), Format$(para.tasageneral * 100, "0000"), "M", "0", "0", " ", " ", " ", "0", "0")


td = "Cta. Cte. Nro. " & Format$(denominACION.ItemData(denominACION.ListIndex), "00000")
mp = Format$(Val(T_TOTAL) * 100, "00000000")
dp = "T"
If rk Then
      rk = epson1.SendInvoicePayment(td, Format$(Val(t_diferencia) * 100, "00000000"), "T")
 Else
      Call verificaerrfiscal(epson1.FiscalStatus, epson1.PrinterStatus)
End If
  
 'subtotal para obtener el importe neto, iva y total impreso en la factura
espere.ProgressBar1.Value = 4
espere.Label1 = "Espere... Cerrando Comprobante Fiscal"

 If rk Then
    rk = epson1.GetInvoiceSubtotal("N", "xx")
 Else
   Call verificaerrfiscal(epson1.FiscalStatus, epson1.PrinterStatus)
 End If
 
 If rk Then
      t_subtotalnc = Format$(Val(epson1.AnswerField_10) / 100, "######0.00")
      t_ivanc = Format$(Val(epson1.AnswerField_6) / 100, "####0.00")
      t_totalnc = Format$(Val(epson1.AnswerField_5) / 100, "######0.00")
 Else
     Call verificaerrfiscal(epson1.FiscalStatus, epson1.PrinterStatus)
 End If
 
 
  If rk Then rk = epson1.CloseInvoice(tipocompfz, letraz, " ")
   
  If rk Then
     t_numnc = epson1.AnswerField_3
  Else
     Call verificaerrfiscal(epson1.FiscalStatus, epson1.PrinterStatus)
  End If
  
  imprime_NCf2 = rk
    
 Exit Function
errf:
 MsgBox ("Error al comunicarse con el impresor fiscal. Verifique que esta encendido y reintente")
 Exit Function
   
End Function




Sub grabanc()
  'On Error GoTo ERRORGRABA
  numintz = saca_ultnumero_int_comp("V")
  t_numintnc = numintz
  Set cl_compvta = New comprobantes_venta
  cl_compvta.sucursal = Val(c_sucursal)
  cl_compvta.actual (3)
  cl_compvta.letra = vta_clientes.t_letrafact
  cl_compvta.numcomp = Val(t_numnc)
  abreviatura = cl_compvta.abreviatura
  ubicacionctacte = cl_compvta.ctacte
  ep = "S"
  cp = Format$(sucursal, "0000") & "-" & Format$(t_numop, "00000000")
  
  contado = "N"
  ssi = Val(t_totalnc)
  moneda = "P"
      
      Set rs = New ADODB.Recordset
      q = "select * from g8 where [id_actividad] = " & 1
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
       codact = rs("id_actividad")
       alicuotaib = rs("alicuota_ib")
       cuentaact = rs("cuenta_contable_venta")
      Else
       codact = 0
       alicuotaib = 0
       cuentaact = para.cuenta_ventas
      End If
      Set rs = Nothing
      
        
      
              
      tiporespiva = vta_clientes.c_iva.ItemData(vta_clientes.c_iva.ListIndex)
       
      idcli = denominACION.ItemData(denominACION.ListIndex)
      
      Set cl_cli = New Clientes
      cl_cli.carga (idcli)
      
      
      If Check3 Then
        T2 = 0
      Else
        T2 = Val(t_totalnc) / Val(fdolar)
      End If
      
      If Check1 Then  'solo dolares
        t1 = 0
      End If
      
      
      
      cn1.BeginTrans
       
       
       QUERY = "INSERT INTO vta_02([num_int], [sucursal], [num_comp], [letra], [id_tipocomp], [id_cliente], [fecha], [id_usuario], [subtotal], [impuestos], [iva], [total]," & _
"[estado], [id_cuenta], [stock], [cta_cte], [grabado], [estado_pago], [recibo_Pago], [observaciones], [cotizacion_dolar], [total_otra_moneda], [moneda], [id_vendedor], " & _
" [VENTA], [CONTADO], [perc_ib], [perc_gan], [perc_iva] , [id_actividad], [alicuota_ib], [alicuota_perc_iva], [canje_cereal], [fecha_vto], [total_bultos],  [valor_declarado], " & _
" [transporte], [direccion_transp], [cuit_transp], [perc_ss], [sucursal_ingreso], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02], [chofer02], [dominio02], " & _
" [dominio_acoplado02], [SALDO_IMPAGO02], [num_z], [cae], [cae_vence], [tipo_op])"




QUERY = QUERY & " VALUES (" & numintz & ", " & Val(c_sucursal) & ", " & Val(t_numnc) & ", '" & vta_clientes.t_letrafact & "', 3" & _
", " & vta_clientes.t_id & ", '" & t_fecha & "', " & para.id_usuario & ", " & Val(t_subtotalnc) & ", 0, " & Val(t_ivanc) & ", " & Val(t_totalnc) & _
", 'A', " & cuentaact & ", '" & cl_compvta.STOCK & "', '" & cl_compvta.ctacte & "', '" & cl_compvta.grabado & "', '" & ep & "', '" & cp & "', 'Dto'" & _
", " & Val(fdolar) & ", " & T2 & ", '" & moneda & "', 1, '" & cl_compvta.venta & "', '" & contado & "', 0, " & _
"0, 0, " & codact & ", 0, 0, 0, '" & t_fecha & "', 0, 0, ' ', ' ', ' ', 0, " & Val(c_sucursal) & _
", '" & Left$(vta_clientes.t_cli, 50) & "', '" & Left$(vta_clientes.t_direccion, 50) & "', '" & Left$(vta_clientes.t_cuit, 20) & "', '" & Left$(vta_clientes.t_localidad, 50) & _
"', " & cl_cli.idtipoiva & ", ' ', ' ', ' ', " & ssi & ", " & para.z_actual & ", 'u2', '01/01/2018', 2)"


cn1.Execute QUERY
        
tib = para.tasaib
QUERY = "INSERT INTO vta_03([num_int], [RENGLON], [id_producto], [descripcion], [cantidad], [pu], [importe], [tasaiva], [impuesto], [costo], [cantidad_original], [tunidad], [pu_final], [tasaib])"
QUERY = QUERY & " VALUES (" & numintz & ", 1, 1, 'Descuento', 1, " & Val(t_subtotalnc) & ", " & Val(t_subtotalnc) & ", " & para.tasageneral & ", 0, 0, 1, 'U', " & Val(t_totalnc) & ", " & tib & ")"

cn1.Execute QUERY
      
        
      
      'actualizo tasa de iva
If cl_compvta.grabado <> "N" Then
          cuentaact = para.cuenta_ventas
          QUERY = "INSERT INTO vta_09([num_int], [tasa_iva], [iva], [neto], [tipo_iva], [id_cuenta09])"
          QUERY = QUERY & " VALUES (" & numintz & ", " & para.tasageneral & ", " & Val(t_ivanc) & ", " & Val(t_subtotalnc) & ", " & cl_cli.idtipoiva & ", " & cuentaact & ")"
          cn1.Execute QUERY
End If
     
      

    'contabilidad
If Generaasientosauto Then
  If cl_compvta.contabilidad <> "N" Then
      numintcgr = saca_ultnumero_int_comp("G")

        u1 = cl_compvta.contabilidad
          
         If u1 = "D" Then
           u2 = "H"
         Else
           u2 = "D"
         End If
         
         
           tot = Val(T_TOTAL)
           m = 1
         
         QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Ventas] " & cl_compvta.abreviatura & " " & t_letra & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & "', 'V', " & numint & ", " & tot & ", " & tot & ", " & para.id_usuario & ", '" & Left$(RTrim$(c_prov), 50) & "')"
         cn1.Execute QUERY
      
         
         
           cta = para.cuenta_deudores
           ic = 1
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cta & ", '" & u1 & "', " & tot & ", '" & dcta & "')"
           cn1.Execute QUERY
           ic = ic + 1
         
         
         
         If Val(t_iva) > 0 And cl_compvta.grabado <> "N" Then
           'cuenta perc
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_iva_ventas & ", '" & u2 & "', " & Format(Val(t_iva) * m, "#####0.00") & ", 'IVA')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
         
         'contrapartida
         
         If cl_compvta.grabado <> "N" Then
           importe = Val(t_subtotal) * m
         Else
           importe = Val(T_TOTAL) * m
         End If
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cuentaact & ", '" & u2 & "', " & Format(importe, "######0.00") & ", '" & "Ventas" & "')"
         cn1.Execute QUERY
         ic = ic + 1
      
      
      End If
      
      
            
     End If
     
      

      
      
      
      
      
     QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
     QUERY = QUERY & " VALUES ('Emitir Factura/NC/ND NI:" & numintz & "', " & para.id_usuario & ", 'V', " & numintz & ", '" & Now & "', '[3] " & cl_cli.letra & " " & Format$(Val(t_sucnc), "0000") & "-" & Format$(Val(t_numnc), "00000000") & "', 12, " & cl_cli.id & ")"
  
     cn1.Execute QUERY

      
      
      cn1.CommitTrans
      
      
      
      
     
      
      
      
      Set cl_cli = Nothing
      Set rs = Nothing
      
      cl_compvta.web (numintz)
      
      Set cl_compvta = Nothing
      
Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos y vuelva a repetir la operacion")
  

End Sub
    
  
