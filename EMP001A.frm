VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form emp_ABM_emp 
   BackColor       =   &H00E0E0E0&
   Caption         =   "EMPLEADOS"
   ClientHeight    =   9045
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      Height          =   1095
      Left            =   7200
      TabIndex        =   10
      Top             =   0
      Width           =   4575
      Begin VB.ComboBox C_vend 
         Height          =   315
         ItemData        =   "EMP001A.frx":0000
         Left            =   1560
         List            =   "EMP001A.frx":000D
         TabIndex        =   15
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
      Begin VB.Label Label3 
         Caption         =   "Estado:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Apellido y Nombre:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton Command5 
         Caption         =   "&Enviar Correo"
         Height          =   735
         Left            =   5400
         Picture         =   "EMP001A.frx":0027
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Listar"
         Height          =   735
         Left            =   4080
         Picture         =   "EMP001A.frx":0331
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Borrar"
         Height          =   735
         Left            =   2760
         Picture         =   "EMP001A.frx":063B
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   735
         Left            =   1440
         Picture         =   "EMP001A.frx":0945
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Agregar"
         Height          =   735
         Left            =   120
         Picture         =   "EMP001A.frx":0C4F
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
         Picture         =   "EMP001A.frx":0F59
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
         Picture         =   "EMP001A.frx":17DB
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
      Bindings        =   "EMP001A.frx":205D
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   10695
      _ExtentX        =   18865
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
      Caption         =   "EMPLEADOS"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "id_LEGAJO"
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
         DataField       =   "denominacion"
         Caption         =   "Apellido y Nombre"
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
         DataField       =   "estado"
         Caption         =   "estado"
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
         DataField       =   "num_cuenta_banco"
         Caption         =   "Cuenta Bancaria"
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
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8790
      Width           =   11790
      _ExtentX        =   20796
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
            TextSave        =   "17:16"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "emp_ABM_emp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer

Private Sub btnacepta_Click()
Call limpia
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub


Private Sub c_vend_LostFocus()
If c_vend.ListIndex < 0 Then
   c_vend.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
If para.id_grupo_modulo_actual >= 5 Then
 emp_abm_emp1!t_funcion = "A"
 emp_abm_emp1.Show
Else
 Call sinpermisos
End If
End Sub

Private Sub Command2_Click()
On Error GoTo e1
If DataGrid1.Bookmark > 0 Then
 If para.id_grupo_modulo_actual >= 5 Then
  If Val(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark)) > 0 Then
   emp_abm_emp1!t_funcion = "M"
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
q = "select * from emp_01 where [id_legajo] = " & Val(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark))
rs.Open q, cn1
 emp_abm_emp1!t_id = rs("id_legajo")
 emp_abm_emp1!t_descripcion = rs("denominacion")
 emp_abm_emp1!t_direccion = rs("estado")
 emp_abm_emp1!t_te = rs("num_cuenta_banco")
 emp_abm_emp1.Show

Set rs = Nothing

Exit Sub
ERROR1:
  MsgBox ("Error al Cargar Empleados. Proc.: LLENACAMPOS")
  Exit Sub
End Sub

Private Sub Command3_Click()
On Error GoTo e1
If DataGrid1.Bookmark > 0 And Val(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark)) > 1 Then
 If para.id_grupo_modulo_actual >= 7 Then
  If Val(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark)) > 1 Then
   emp_abm_emp1!t_funcion = "B"
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
Call imprime
End Sub

Sub imprime()
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 0
    c(1) = 1
    c(2) = 2
    c(3) = 3
    For i = 4 To 14
      c(i) = -1
    Next i
    Call imprimedatagrid(DataGrid1, c(), "LISTADO DE EMPLEADOS", "", "ESTADO: " & c_vend, " ", 60, 7, True, False, "H")
  End If



End Sub




Private Sub DataGrid1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime "

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


If KeyCode = vbKeyF7 Then
 If DataGrid1.Bookmark > 0 Then
   Call imprime
 End If
End If
End Sub

Private Sub DataGrid1_LostFocus()
Call barraesag(Me)
End Sub

Private Sub Form_Activate()
Call limpia
DataGrid1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     gen_tools.Show
End Select
End Sub
Sub limpia()
Dim q As String
q = "select * from emp_01 "
c = " where "
If t_prov <> "" Then
 q = q & c & " [denominacion] like '%" & t_prov & "%'"
 c = " and "
End If

If c_vend.ListIndex > 0 Then
 q = q & c & " [estado] ='" & Mid$(c_vend, 1, 1) & "'"
 c = " and "
End If


Call conectaradodc(Adodc1, q, cn1)
DataGrid1.Refresh
Call INICIALIZA2(emp_abm_emp1)
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Load emp_abm_emp1
c_vend.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload emp_abm_emp1
End Sub

Private Sub t_prov_GotFocus()
t_prov = ""
End Sub


