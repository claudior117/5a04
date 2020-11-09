VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form gen_agenda 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "AGENDA DE EVENTOS DIARIOS"
   ClientHeight    =   4020
   ClientLeft      =   915
   ClientTop       =   3075
   ClientWidth     =   7560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7560
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   6120
      TabIndex        =   7
      Top             =   3000
      Width           =   1335
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   720
         Picture         =   "arch014A.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salir sin Modificar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton btnacepta 
         Height          =   495
         Left            =   120
         Picture         =   "arch014A.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   3000
      Width           =   5535
      Begin VB.ComboBox c_usuario 
         Height          =   315
         Left            =   3480
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox t_fecha 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Usuario:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2280
      Top             =   2760
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
      UserName        =   ""
      Password        =   ""
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
      Bindings        =   "arch014A.frx":1104
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   49152
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   24
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
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "AGENDA"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "fecha_vto"
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
      BeginProperty Column01 
         DataField       =   "descripcion"
         Caption         =   "Evento"
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
      BeginProperty Column02 
         DataField       =   "observaciones"
         Caption         =   "Observaciones"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "usuario"
         Caption         =   "Usuario"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Id_evento"
         Caption         =   "Id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
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
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3765
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   21167
            MinWidth        =   21167
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "gen_agenda"
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


Private Sub c_usuario_LostFocus()
If c_usuario.ListIndex < 0 Then
  c_usuario.ListIndex = buscaindice(c_usuario, para.id_usuario)
End If
End Sub



Sub LLENACAMPOS()
'On Error GoTo ERROR1
Set rs = New ADODB.Recordset
q = "select * from g7 where [id_evento] = " & Val(DataGrid1.Columns(4).CellValue(DataGrid1.Bookmark))
rs.Open q, cn1
If rs("id_usuario") = para.id_usuario Then
 gen_agenda1!t_id = rs("id_evento")
 gen_agenda1!t_descripcion = rs("descripcion")
 gen_agenda1!t_obs = rs("observaciones")
 gen_agenda1!t_fecha = rs("fecha_vto")
 gen_agenda1.Show
Else
  MsgBox ("No se puede modificar o borrar un evento creado por otro usuario")
End If

Set rs = Nothing

Exit Sub
ERROR1:
  MsgBox ("Error al Cargar Eventos. Proc.: LLENACAMPOS")
  Exit Sub
End Sub






Private Sub DataGrid1_GotFocus()
StatusBar1.Panels(1).Text = "[F1] Agrega Evento  - [F3] Reprograma Evento - [F8] Borra Evento  - [ESC] Cierra Agenda"
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
   Case Is = vbKeyF1
        gen_agenda1!t_funcion = "A"
        gen_agenda1.Show
   Case Is = vbKeyF3
        On Error GoTo e1
        If DataGrid1.Bookmark > 0 Then
          If Val(DataGrid1.Columns(4).CellValue(DataGrid1.Bookmark)) > 0 Then
            gen_agenda1!t_funcion = "M"
            Call LLENACAMPOS
           End If
        End If
   Case Is = vbKeyF8
        On Error GoTo e1
        If DataGrid1.Bookmark > 0 Then
          If Val(DataGrid1.Columns(4).CellValue(DataGrid1.Bookmark)) > 0 Then
            J = MsgBox("Confirma eliminar el evento", 4)
            If J = 6 Then
                  QUERY = "DELETE FROM g7 WHERE [id_evento] = " & Val(DataGrid1.Columns(4).CellValue(DataGrid1.Bookmark))
                  cn1.BeginTrans
                    cn1.Execute QUERY
                  cn1.CommitTrans
                  Call limpia
                  
            End If
          End If
        End If
End Select

Exit Sub
e1:
 Exit Sub

End Sub
Sub limpia()
Dim q As String
q = "select * from g7, g1 where g7.[id_usuario] = g1.[id_usuario] "
If t_fecha <> "" Then
   q = q & " and datevalue([fecha_vto]) >= datevalue('" & t_fecha & "')"
End If

If c_usuario.ListIndex > 0 Then
   q = q & " and g7.[id_usuario] = " & c_usuario.ItemData(c_usuario.ListIndex)
End If
q = q & " order by [fecha_vto], [id_evento]"
Call conectaradodc(Adodc1, q, cn1)
DataGrid1.Refresh
Call INICIALIZA2(gen_agenda1)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Unload Me
End If
  
End Sub

Private Sub Form_Load()
Call carga_usuarios(c_usuario)
c_usuario.AddItem "<Todos>", 0
c_usuario.ListIndex = buscaindice(c_usuario, para.id_usuario)
Load gen_agenda1
t_fecha = Format$(Now, "dd/mm/yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload gen_agenda1
End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
     t_fecha = ""
  Else
     t_fecha = Format$(t_fecha, "dd/mm/yyyy")
  End If
End If
End Sub
