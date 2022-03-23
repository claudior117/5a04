VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ver_PROD_oc 
   BackColor       =   &H00E0E0E0&
   Caption         =   "VER PRODUCTOS PENDIENTES EN ORDEN DE COMPRA"
   ClientHeight    =   8655
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   12105
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      Height          =   1575
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   11535
      Begin VB.TextBox t_fecha2 
         Height          =   285
         Left            =   8400
         MaxLength       =   10
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox c_obra 
         Height          =   315
         Left            =   1440
         TabIndex        =   14
         Top             =   960
         Width           =   4815
      End
      Begin VB.ComboBox c_estado 
         Height          =   315
         ItemData        =   "Proc002.frx":0000
         Left            =   8400
         List            =   "Proc002.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   960
         Width           =   2535
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox t_fecha 
         Height          =   285
         Left            =   8400
         MaxLength       =   10
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox c_prod 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Hasta:"
         Height          =   255
         Left            =   6600
         TabIndex        =   17
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Obra"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Estado Recepcion:"
         Height          =   255
         Left            =   6600
         TabIndex        =   13
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Proveedor"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Desde:"
         Height          =   255
         Left            =   6600
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Producto"
         Height          =   495
         Left            =   120
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
         Picture         =   "Proc002.frx":003A
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
         Picture         =   "Proc002.frx":08BC
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
      Bindings        =   "Proc002.frx":113E
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   10398
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
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
      Caption         =   "PRODUCTOS EN ORDENES DE COMPRA "
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "num_comprobante"
         Caption         =   "Num. OC"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "00000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "a5.Fecha"
         Caption         =   "Fecha"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "a6.id_producto"
         Caption         =   "Id."
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
      BeginProperty Column04 
         DataField       =   "a6.Detalle"
         Caption         =   "Producto"
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
         DataField       =   "total_oc"
         Caption         =   "En O.C"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "total_recibido"
         Caption         =   "Recibido"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "total_pedido"
         Caption         =   "Pedido"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "a5.num_int"
         Caption         =   "Num.Int."
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
      BeginProperty Column09 
         DataField       =   "id_proveedor"
         Caption         =   "id. proveedor"
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
      BeginProperty Column10 
         DataField       =   "renglon"
         Caption         =   "renglon"
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
      BeginProperty Column11 
         DataField       =   "pu"
         Caption         =   "P.U.en OC"
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
      BeginProperty Column12 
         DataField       =   "num_int_item"
         Caption         =   "Referencia"
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
      BeginProperty Column13 
         DataField       =   "descripcion"
         Caption         =   "Obra"
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
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column08 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column09 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column12 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column13 
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8400
      Width           =   12105
      _ExtentX        =   21352
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
            TextSave        =   "11/03/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "08:47 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "ver_PROD_oc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer

Private Sub btnacepta_Click()
Call limpia
DataGrid1.SetFocus

End Sub

Private Sub btnsale_Click()
Unload Me
End Sub




Private Sub DataGrid1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F9] Recepcion - "

End Sub
Sub ARMARECEP()

ABM_COMP_COMPRA.Show

i = 1
'FIXIT: Declare 'varBmk' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Dim varBmk As Variant
Set cl_prod = New productos
For Each varBmk In DataGrid1.SelBookmarks
  
 Adodc1.Recordset.Bookmark = varBmk
   ip = Format$(DataGrid1.Columns(3).CellValue(DataGrid1.Bookmark), "00000")
   cl_prod.cargar (ip)
   d = DataGrid1.Columns(4).CellValue(DataGrid1.Bookmark)
   cu = Format$(Val(DataGrid1.Columns(5).CellValue(DataGrid1.Bookmark)) - Val(DataGrid1.Columns(6).CellValue(DataGrid1.Bookmark)), "#####0.00")
   pu = Format$(Val(DataGrid1.Columns(11).CellValue(DataGrid1.Bookmark)), "#####0.00")
   ti = Format$(para.tasageneral, "##0.00")
   i = Format$(Val(cu) * Val(pu), "######0.00")
   r = ABM_COMP_COMPRA.msf1.Rows
   nir = Val(DataGrid1.Columns(12).CellValue(DataGrid1.Bookmark))
   ABM_COMP_COMPRA.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & cu & Chr(9) & pu & Chr(9) & ti & Chr(9) & "" & Chr$(9) & i & Chr(9) & nir
Next
ABM_COMP_COMPRA.sacatotales2
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


Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF9 Then
     J = MsgBox("Confirma recepcion de materiales (Armar Comprobante de Compra)", 4)
     If J = 6 Then
       Call ARMARECEP
     End If
End If

If KeyCode = vbKeyF7 Then
   Dim c(15) As Double
   J = MsgBox("Prepare Impresora y Confirme)", 4)
   If J = 6 Then
       c(0) = 0
       c(1) = 1
       c(2) = 2
       c(3) = 4
       c(4) = 5
       c(5) = 6
       c(6) = 7
       c(7) = 13
       
       For i = 8 To 14
         c(i) = -1
       Next i
        Call imprimedatagrid(DataGrid1, c, Space$(35) & "Productos Pendientes en orden de Compra", "     Fecha desde........:" & t_fecha, "     Estado recepcion...:" & c_estado, "     Obra...............:", 45, 9, True, True, "H")
     End If
End If


End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Load cc_detalle
    cc_detalle.t_idprov = DataGrid1.Columns(8).CellValue(DataGrid1.Bookmark)
    cc_detalle.t_prov = DataGrid1.Columns(2).CellValue(DataGrid1.Bookmark)
    cc_detalle.t_sucursal = glo.sucursal
    cc_detalle.t_letra = "O"
    cc_detalle.t_numcomp = DataGrid1.Columns(1).CellValue(DataGrid1.Bookmark)
    cc_detalle.t_tipocomp = 65
    cc_detalle.t_numint = DataGrid1.Columns(8).CellValue(DataGrid1.Bookmark)
    cc_detalle.Show
End If
End Sub

Function busca_prod_oc(cp As Long) As Long
  i = 1
  r = 0
  While i <= ABM_COMP_COMPRA.msf1.Rows - 1
     If cp = Val(ABM_COMP_COMPRA.msf1.TextMatrix(i, 1)) Then
        r = i
        i = ABM_COMP_COMPRA.msf1.Rows
     End If
     i = i + 1
  Wend
  busca_prod_oc = r
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub
Sub limpia()
Dim q As String
q = "select * from a6, a5,a1, pro_04, a4 where [id_tipocomp] = 65 and a5.[NUM_INT] = a6.[NUM_INT] and a5.[id_proveedor] = a1.[id_proveedor] and a6.[num_int_item] = pro_04.[num_referencia] and a6.[id_obra] = a4.[id_obra]"
c = " and "
Select Case c_prod.ListIndex
 Case Is > 0
    q = q & c & " a6.[id_producto] = " & c_prod.ItemData(c_prod.ListIndex)
 Case Is < 0
  If c_prod <> "" Then
    q = q & c & " pro_04.[detalle] like '%" & c_prod & "%'"
    c = " and "
  End If
End Select

If t_fecha <> "" Then
   q = q & c & " datevalue(a5.[fecha]) >= datevalue('" & t_fecha & "')"
   c = " and "
 End If


If t_fecha2 <> "" Then
   q = q & c & " datevalue(a5.[fecha]) <= datevalue('" & t_fecha2 & "')"
   c = " and "
 End If


If c_prov.ListIndex > 0 Then
     q = q & c & " a5.[id_proveedor] = " & c_prov.ItemData(c_prov.ListIndex)
End If

If C_OBRA.ListIndex > 0 Then
     q = q & c & " a6.[id_obra] = " & C_OBRA.ItemData(C_OBRA.ListIndex)
End If


If c_estado.ListIndex > 0 Then
  Select Case c_estado.ListIndex
   Case Is = 1
     q = q & " and [total_recibido] >= [total_pedido]"
   Case Is = 2
     q = q & " and [total_recibido] < [total_pedido]"
  End Select
End If

Call conectaradodc(Adodc1, q, cn1)
DataGrid1.Refresh
End Sub

Private Sub Form_Load()

'Call carga_productos(c_prod)
c_prod.AddItem "<Todos>", 0
c_prod.ListIndex = 0

Call carga_proveedores(c_prov)
c_prov.AddItem "<Todos>", 0
c_prov.ListIndex = 0

Call carga_obras(C_OBRA, "A")
C_OBRA.AddItem "<Todos>", 0
C_OBRA.ListIndex = 0

c_estado.ListIndex = 2

Call barraesag(Me)


End Sub

Private Sub t_fecha_LostFocus()
Call solofecha(t_fecha)
End Sub

Private Sub t_fecha2_LostFocus()
Call solofecha(t_fecha2)
End Sub
