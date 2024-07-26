VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form vta_ABM_rempredef 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REMITOS PREDEFINIDOS"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7620
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
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
         Picture         =   "vta074A.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Borrar"
         Height          =   735
         Left            =   2760
         Picture         =   "vta074A.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   735
         Left            =   1440
         Picture         =   "vta074A.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Agregar"
         Height          =   735
         Left            =   120
         Picture         =   "vta074A.frx":091E
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
      Left            =   5640
      TabIndex        =   2
      Top             =   6600
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta074A.frx":0C28
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
         Picture         =   "vta074A.frx":14AA
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
      Left            =   600
      Top             =   6840
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
      Bindings        =   "vta074A.frx":1D2C
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   6975
      _ExtentX        =   12303
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
      Caption         =   "REMITOS PREDEFINIDOS"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "id_rempredef"
         Caption         =   "Id"
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
      BeginProperty Column01 
         DataField       =   "denominacion"
         Caption         =   "Identificacion"
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
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   7845
      Width           =   7620
      _ExtentX        =   13441
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
            TextSave        =   "25/07/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:47 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_ABM_rempredef"
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
If para.id_grupo_modulo_actual >= 6 Then   'ADMINISTRADOR Y VENTAS
 vta_abm_rempredef1!t_funcion = "A"
 vta_abm_rempredef1.Show
Else
 Call sinpermisos
End If
End Sub

Private Sub Command2_Click()
On Error GoTo e1
If DataGrid1.Bookmark > 0 Then
 If para.id_grupo_modulo_actual >= 5 Then
  If Val(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark)) > 1 Then
   vta_abm_rempredef1!t_funcion = "M"
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
q = "select * from vta_17 where [id_rempredef] = " & Val(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark))
rs.Open q, cn1
 vta_abm_rempredef1!t_id = rs("id_rempredef")
 vta_abm_rempredef1!t_descripcion = rs("denominacion")
 
 
 vta_abm_rempredef1.Show

Set rs = Nothing

Exit Sub
ERROR1:
  MsgBox ("Error al Cargar Vendedores. Proc.: LLENACAMPOS")
  Exit Sub
End Sub

Private Sub Command3_Click()
On Error GoTo e1
If DataGrid1.Bookmark > 0 And Val(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark)) > 1 Then
 If para.id_grupo_modulo_actual >= 7 Then
  If Val(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark)) > 1 Then
   vta_abm_rempredef1!t_funcion = "B"
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
'Call ejecutareporte(Adodc1, arch001)
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
q = "select * from vta_017 "
Call conectaradodc(Adodc1, q, cn1)
DataGrid1.Refresh
Call INICIALIZA2(vta_abm_rempredef1)
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Load vta_abm_rempredef1

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload vta_abm_rempredef1
End Sub

