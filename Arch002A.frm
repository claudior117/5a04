VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ABM_PROD 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PRODUCTOS"
   ClientHeight    =   8790
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12135
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   12135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      Height          =   1095
      Left            =   6840
      TabIndex        =   10
      Top             =   0
      Width           =   4575
      Begin VB.TextBox t_localidad 
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox t_prov 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Producto"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Proveedor"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton Command4 
         Caption         =   "&Listar"
         Height          =   735
         Left            =   4080
         Picture         =   "Arch002A.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Borrar"
         Height          =   735
         Left            =   2760
         Picture         =   "Arch002A.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   735
         Left            =   1440
         Picture         =   "Arch002A.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Agregar"
         Height          =   735
         Left            =   120
         Picture         =   "Arch002A.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   2
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "Arch002A.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "Arch002A.frx":14AA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   720
      Top             =   7800
      Visible         =   0   'False
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   661
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "claudio"
      Password        =   "0969"
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Arch002A.frx":1D2C
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      Caption         =   "PRODUCTOS"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "id_producto"
         Caption         =   "Id."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "00000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "descripcion"
         Caption         =   "Detalle"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0000000000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "unidad"
         Caption         =   "Unidad"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "envase"
         Caption         =   "Envase"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "denominacion"
         Caption         =   "Proveedor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "talle"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "color"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "medida"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   5
         SizeMode        =   1
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8535
      Width           =   12135
      _ExtentX        =   21405
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
            TextSave        =   "19/08/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:17 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "ABM_PROD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer

Private Sub btnsale_Click()
Unload Me
End Sub


Private Sub Command1_Click()
Call nivel_acceso(5)
If para.id_grupo_modulo_actual >= 5 Then
 Call INICIALIZA2(abm_prod1)
 abm_prod1!t_funcion = "A"
 abm_prod1.Show
Else
 Call sinpermisos
End If
End Sub

Private Sub Command2_Click()
On Error GoTo e1
If DataGrid1.Bookmark > 0 Then
 Call nivel_acceso(5)
 If para.id_grupo_modulo_actual >= 5 Then
  If Val(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark)) > 1 Then
   abm_prod1!t_funcion = "M"
   Call LLENACAMPOS
  End If
 Else
  Call sinpermisos
 End If
End If

Exit Sub
e1:
 Exit Sub
End Sub

Sub LLENACAMPOS()
On Error GoTo ERROR1
Set rs = New ADODB.Recordset
q = "select * from a2, a1, g5 where [id_producto] = " & Val(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark)) & " and a2.[id_proveedor] = a1.[id_proveedor] and a2.[id_unidad] = g5.[id_unidad]"
rs.Open q, cn1
 abm_prod1!t_id = rs("id_producto")
 abm_prod1!t_descripcion = rs("descripcion")
 abm_prod1!c_prov.ListIndex = buscaindice(abm_prod1!c_prov, rs("a2.id_proveedor"))
 abm_prod1!c_unidad.ListIndex = buscaindice(abm_prod1!c_unidad, rs("a2.id_unidad"))
 abm_prod1!t_envase = rs("envase")
 abm_prod1!c_grupo.ListIndex = buscaindice(abm_prod1!c_grupo, rs("id_grupo"))
 abm_prod1!t_pu = rs("pu")
 abm_prod1!c_iva.ListIndex = buscaindice(abm_prod1!c_iva, rs("cod_tasaiva"))
 abm_prod1!t_stockminimo = rs("stock_minimo")
 abm_prod1!c_marca.ListIndex = buscaindice(abm_prod1!c_marca, rs("id_marca"))
 abm_prod1!c_depto.ListIndex = buscaindice(abm_prod1!c_depto, rs("id_departamento"))
 abm_prod1!t_utilidad = rs("porc_utilidad")
 abm_prod1!t_costo = rs("costoreal")
 abm_prod1!t_fletecompra = rs("flete_compra")
 abm_prod1!t_dtocompra = rs("dto_compra")
 abm_prod1!t_codbarra = rs("cod_barra")
 abm_prod1!t_final = rs("precio_final")
 abm_prod1!t_tasaimpint = rs("tasa_imp_interno")
 abm_prod1!t_tipo = rs("tipo_producto")
 abm_prod1!t_moneda = rs("moneda")
 abm_prod1!t_impuesto = rs("impuesto")
 abm_prod1!t_observaciones = rs("observaciones")
 abm_prod1!t_preciocompra = rs("precio_ult_compra")
 abm_prod1!t_abreviatura = rs("texto_central")
 abm_prod1!c_tasaib.ListIndex = buscaindice(abm_prod1!c_tasaib, rs("id_tasaib"))
 abm_prod1!t_talle = rs("talle")
 abm_prod1!t_color = rs("color")
 abm_prod1!t_medida = rs("medida")


 
 abm_prod1.Show

Set rs = Nothing

Exit Sub
ERROR1:
  MsgBox ("Error al Cargar Productos. Proc.: LLENACAMPOS")
  Exit Sub
End Sub

Private Sub Command3_Click()
On Error GoTo e1
If DataGrid1.Bookmark > 0 And Val(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark)) > 1 Then
 Call nivel_acceso(5)
 If para.id_grupo_modulo_actual >= 7 Then
  If Val(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark)) > 1 Then
   abm_prod1!t_funcion = "B"
   Call LLENACAMPOS
  End If
 Else
  Call sinpermisos
 End If
End If

Exit Sub
e1:
 Exit Sub
End Sub

Private Sub Command4_Click()
'Call ejecutareporte(Adodc1, arch002)
End Sub





Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
Set rs = Adodc1.Recordset
gcolumna = ColIndex
If rs.Sort <> DataGrid1.Columns(ColIndex).DataField & " ASC" Then
   rs.Sort = DataGrid1.Columns(ColIndex).DataField & " ASC"
Else
   rs.Sort = DataGrid1.Columns(ColIndex).DataField & " DESC"
End If
Set rs = Nothing

End Sub


Private Sub Form_Activate()
Call limpia
DataGrid1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub
Sub limpia()
Dim q As String
q = "select * from a2, g5, a1 where a2.[id_unidad] = g5.[id_unidad] and a2.[id_proveedor] = a1.[id_proveedor] "
c = " and "
If t_prov <> "" Then
 q = q & c & " [denominacion] like '%" & t_prov & "%'"
 c = " and "
End If
If t_localidad <> "" Then
 q = q & c & " a2.[descripcion] like '%" & t_localidad & "%'"
 c = " and "
End If


Call conectaradodc(Adodc1, q, cn1)
DataGrid1.Refresh
Call INICIALIZA2(abm_prod1)
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Load abm_prod1

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload abm_prod1
End Sub

Private Sub t_prov_GotFocus()
t_prov = ""
End Sub

Private Sub t_prov_LostFocus()
Call limpia
End Sub
Private Sub t_localidad_GotFocus()
t_localidad = ""
End Sub

Private Sub t_localidad_LostFocus()
Call limpia
End Sub

