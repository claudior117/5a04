VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form com_config_comp 
   BackColor       =   &H00E0E0E0&
   Caption         =   "CONFIGURA COMPROBANTES DE COMPRA"
   ClientHeight    =   8805
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   12240
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton Command4 
         Caption         =   "&Listar"
         Height          =   735
         Left            =   1440
         Picture         =   "com002.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   735
         Left            =   120
         Picture         =   "com002.frx":030A
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
         Picture         =   "com002.frx":0614
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
         Picture         =   "com002.frx":0E96
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
      Bindings        =   "com002.frx":1718
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   11655
      _ExtentX        =   20558
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
      Caption         =   "COMPROBANTES DE COMPRA"
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "id_tipo_comp"
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
         Caption         =   "Comprobante"
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
         DataField       =   "abreviatura"
         Caption         =   "Abrevia"
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
         DataField       =   "ult_num"
         Caption         =   "Ult.. Num"
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
         DataField       =   "copias"
         Caption         =   "Copias "
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
         DataField       =   "stock"
         Caption         =   "Stock"
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
         DataField       =   "Iva"
         Caption         =   "Iva"
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
         DataField       =   "Compra"
         Caption         =   "Inf. Compra"
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
      BeginProperty Column08 
         DataField       =   "contabilidad"
         Caption         =   "Contabilidad"
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
         DataField       =   "formato"
         Caption         =   "Formato"
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
         DataField       =   "imprime_desc_extra"
         Caption         =   "Imprime desc. extra"
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
            Alignment       =   1
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8550
      Width           =   12240
      _ExtentX        =   21590
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
            TextSave        =   "09:41"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "com_config_comp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer

Private Sub btnsale_Click()
Unload Me
End Sub



Private Sub c_sucursal_LostFocus()
Call limpia
DataGrid1.SetFocus
End Sub

Private Sub Command1_Click()
Call limpia
DataGrid1.SetFocus
End Sub

Private Sub Command2_Click()
'On Error GoTo e1
If DataGrid1.Bookmark > 0 Then
 If para.id_grupo_modulo_actual >= 8 Then
  If Val(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark)) > 0 Then
   com_config_comp1!t_funcion = "M"
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
'On Error GoTo ERROR1
Set rs = New ADODB.Recordset
q = "select * from g2 where  [id_tipo_comp] = " & Val(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark))
rs.Open q, cn1
 'com_config_comp1!t_sucursal = rs("sucursal")
 com_config_comp1!t_id = rs("id_tipo_comp")
 com_config_comp1!t_descripcion = rs("descripcion")
 com_config_comp1!t_abrevia = rs("abreviatura")
 com_config_comp1!t_ultimoA = rs("ult_num")
 com_config_comp1!t_stock = rs("stock")
 com_config_comp1!t_ctacte = rs("ctacte")
 com_config_comp1!t_iva = rs("iva")
 com_config_comp1!t_venta = rs("compra")
 com_config_comp1!t_contab = rs("contabilidad")
 com_config_comp1!t_copiasA = rs("copias")
 com_config_comp1!t_formato2 = rs("formato")
 com_config_comp1!t_ie = rs("imprime_desc_extra")
 
 com_config_comp1.Show

Set rs = Nothing

Exit Sub
ERROR1:
  MsgBox ("Error al Cargar Comprobantes. Proc.: LLENACAMPOS")
  Exit Sub
End Sub


Private Sub Command4_Click()
If para.id_grupo_modulo_compras > 4 Then
  J = MsgBox("Prepare Impresora y Confirme", 4)
    If J = 6 Then
      Dim c(15) As Double
      c(0) = 0
      c(1) = 1
      c(2) = 2
      c(3) = 3
      c(4) = 4
      c(5) = 10
      c(6) = 11
      c(7) = 12
      c(8) = 13
      c(9) = 14
      c(10) = 18
      For i = 11 To 14
        c(i) = -1
      Next i
      Call imprimedatagrid(DataGrid1, c(), "CONFIGURACION COMPROBANTES DE COMPRA", "", "Sucursal: " & c_sucursal, "", 90, 6, True, False)
     End If
         
  End If

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
q = "select * from G2 ORDER by [id_tipo_comp]"
Call conectaradodc(Adodc1, q, cn1)
DataGrid1.Refresh
Call INICIALIZA2(com_config_comp1)
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Load com_config_comp1

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload com_config_comp1
End Sub

