VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form exp_exporta 
   BackColor       =   &H00E0E0E0&
   Caption         =   "OPERACIONES DE EXPORTACION "
   ClientHeight    =   8670
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12225
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordenados por"
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   7320
      Width           =   4455
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Id."
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Embarque"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      Height          =   1335
      Left            =   7080
      TabIndex        =   9
      Top             =   0
      Width           =   4695
      Begin VB.ComboBox c_estado 
         Height          =   315
         ItemData        =   "EXP001A.frx":0000
         Left            =   1080
         List            =   "EXP001A.frx":000D
         TabIndex        =   19
         Text            =   "Combo1"
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox t_detalle 
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox t_cli 
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000C0&
         Caption         =   "Estado:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000C0&
         Caption         =   "Detalle:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   "Cliente:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton Command4 
         Caption         =   "&Listar"
         Height          =   735
         Left            =   4080
         Picture         =   "EXP001A.frx":0032
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Borrar"
         Height          =   735
         Left            =   2760
         Picture         =   "EXP001A.frx":033C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   735
         Left            =   1440
         Picture         =   "EXP001A.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Agregar"
         Height          =   735
         Left            =   120
         Picture         =   "EXP001A.frx":0950
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   1
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "EXP001A.frx":0C5A
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
         Picture         =   "EXP001A.frx":14DC
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
      Top             =   8415
      Width           =   12225
      _ExtentX        =   21564
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
            TextSave        =   "28/10/2010"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "07:01 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5775
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   10186
      _Version        =   393216
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "exp_exporta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim gquery As String

Private Sub btnacepta_Click()
Call limpia
msf1.SetFocus

End Sub

Private Sub btnsale_Click()
Unload Me
End Sub





Private Sub c_estado_LostFocus()
If c_estado.ListIndex < 0 Then
  c_estado.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
If para.id_grupo_modulo_actual >= 5 Then
 exp_exporta1!t_funcion = "A"
 exp_exporta1.Show
Else
 Call sinpermisos
End If
End Sub

Private Sub Command2_Click()
'On Error GoTo e1
If msf1.Rows > 0 Then
 If para.id_grupo_modulo_actual >= 5 Then
  If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
   exp_exporta1!t_funcion = "M"
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
q = "select * from exp01 where [num_exp] = " & Val(msf1.TextMatrix(msf1.Row, 0))
 rs.Open q, cn1
 exp_exporta1!t_id = rs("num_exp")
 exp_exporta1!t_descripcion = rs("detalle")
 exp_exporta1.c_cli.ListIndex = buscaindice(exp_exporta1.c_cli, rs("id_cliente"))
 exp_exporta1!t_fechap = rs("fecha_embarque")
 exp_exporta1!t_fechaf = rs("fecha_fact")
 exp_exporta1!t_importe = rs("importe")
 exp_exporta1.Show

Set rs = Nothing

Exit Sub
ERROR1:
  MsgBox ("Error al Cargar Operacion de Exportacion. Proc.: LLENACAMPOS")
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
    For i = 2 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "LISTADO DE OPERACIONES DE EXPORTACION", "Cliente: " & t_cli, "Detalle: " & t_detalle, " ", 60, 7, True, False)
      
  End If


End Sub





Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 9
msf1.FixedCols = 1
msf1.SelectionMode = flexSelectionFree
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 500
msf1.ColWidth(1) = 4500
msf1.ColWidth(2) = 500
msf1.ColWidth(3) = 3500
msf1.ColWidth(4) = 1000
msf1.ColWidth(5) = 1000
msf1.ColWidth(6) = 1000
msf1.ColWidth(7) = 1000
msf1.ColWidth(8) = 1000
msf1.TextMatrix(0, 0) = "Id."
msf1.TextMatrix(0, 1) = "Detalle"
msf1.TextMatrix(0, 2) = "Id."
msf1.TextMatrix(0, 3) = "Cliente"
msf1.TextMatrix(0, 4) = "Embarque"
msf1.TextMatrix(0, 5) = "Facturado"
msf1.TextMatrix(0, 6) = "Importe"
msf1.TextMatrix(0, 7) = "Ingresado"
msf1.TextMatrix(0, 8) = "Estado"
For i = 1 To 5
  msf1.ColAlignment(i) = 1 'izq
Next i
msf1.ColAlignment(6) = 9 'der
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
     
End Select
End Sub
Sub limpia()
Dim q As String
Call armagrid
espere.Show
espere.Label1 = "ESPERE [Leyendo Base de Datos]... "
espere.Refresh

q = "select * from exp01"
c = " where "
If t_cli <> "" Then
 q = q & c & " [cliente] like '%" & t_cli & "%'"
 c = " and "
End If

If t_detalle <> "" Then
 q = q & c & " [detalle] like '%" & t_detalle & "%'"
 c = " and "
End If

If c_estado.ListIndex > 0 Then
  q = q & c & " [estado] = '" & Mid$(c_estado, 1, 1) & "'"
End If

If Option2 = True Then
  q = q & " order by [fecha_embarque]"
Else
   q = q & " order by [num_exp]"
End If


Set rs = New ADODB.Recordset
rs.Open q, cn1
c = 0
While Not rs.EOF
 tr = sacareintegro(rs("num_exp"))
 If rs("estado") = "E" Then
   e = "En Proceso"
 Else
   e = "Terminada"
 End If
 
 msf1.AddItem rs("num_exp") & Chr$(9) & rs("detalle") & Chr$(9) & rs("id_cliente") & Chr$(9) & rs("cliente") & Chr$(9) & rs("fecha_embarque") & Chr$(9) & rs("fecha_fact") & Chr$(9) & rs("importe") & Chr$(9) & tr & Chr$(9) & e
 rs.MoveNext
 c = c + 1
Wend
 msf1.AddItem ""
 msf1.AddItem "" & Chr$(9) & "Total de Registros : " & c

Set rs = Nothing
Call INICIALIZA2(exp_exporta1)
Unload espere
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Load exp_exporta1
Option2 = True
Call armagrid
t_detalle = ""
t_cli = ""
c_estado.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload exp_exporta1
End Sub

Private Sub msf1_GotFocus()
StatusBar1.Panels.Item(2) = "[F4] Saca - [F7] Imprime - [F3] Planilla Reintegro - [F5] Cambia estado "

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF3 Then
 If msf1.Rows > 0 Then
   Load exp_prodreintegro
   exp_prodreintegro.c_prov.ListIndex = buscaindice(exp_prodreintegro.c_prov, Val(msf1.TextMatrix(msf1.Row, 2)))
   Call exp_prodreintegro.carga_exportaciones(exp_prodreintegro.c_vend)
   exp_prodreintegro.c_vend.ListIndex = buscaindice(exp_prodreintegro.c_vend, Val(msf1.TextMatrix(msf1.Row, 0)))
   exp_prodreintegro.carga
   exp_prodreintegro.Show
   
   
   
  End If
End If

If KeyCode = vbKeyF5 Then
 If msf1.Rows > 0 Then
   r = msf1.Row
   J = MsgBox("Confirma cambiar estado operacion Nro. " & msf1.TextMatrix(r, 0), 4)
   If J = 6 Then
     If msf1.TextMatrix(r, 8) = "Terminada" Then
       e = "E"
     Else
       e = "T"
     End If
     
     cn1.BeginTrans
     QUERY = "update exp01 set  [estado]= '" & e & "'"
     QUERY = QUERY & " where [num_exp]= " & Val(msf1.TextMatrix(r, 0))
     cn1.Execute QUERY
     cn1.CommitTrans
   
     Call limpia
   
   End If
  End If
End If


If KeyCode = vbKeyF4 Then
 If msf1.Rows > 0 Then
   msf1.RemoveItem msf1.Row
  End If
End If


If KeyCode = vbKeyF7 Then
 If msf1.Rows > 0 Then
   
   Call imprime
    
 End If
End If

End Sub



Private Sub t_cli_GotFocus()
t_cli = ""
End Sub

Private Sub T_detalle_GotFocus()
t_detalle = ""
End Sub
