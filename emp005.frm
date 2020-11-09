VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form emp_emitegastos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REGISTRO DE GASTOS POR EMPLEADO"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10395
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Borrar Mov."
      Height          =   375
      Left            =   8400
      TabIndex        =   20
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   7680
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Funcion"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Totales del Comprobante"
      Height          =   1095
      Left            =   240
      TabIndex        =   14
      Top             =   6840
      Width           =   7935
      Begin VB.TextBox t_observaciones 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   49
         TabIndex        =   4
         Top             =   240
         Width           =   6015
      End
      Begin VB.TextBox t_total 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Total"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5055
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8916
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   240
      TabIndex        =   10
      Top             =   0
      Width           =   9375
      Begin VB.ComboBox c_tipocomp 
         Height          =   315
         ItemData        =   "emp005.frx":0000
         Left            =   2160
         List            =   "emp005.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   5175
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox t_numcomp 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Empleado:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Nro. Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   8400
      TabIndex        =   7
      Top             =   6960
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "emp005.frx":003D
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "emp005.frx":08BF
         Style           =   1  'Graphical
         TabIndex        =   8
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
      TabIndex        =   6
      Top             =   8745
      Width           =   10395
      _ExtentX        =   18336
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
            TextSave        =   "04/02/2020"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "08:57"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "emp_emitegastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Dim cantidadp As Double
Sub limpia()
   Call armagrid
   t_subtotal = ""
   t_nograbado = ""
   t_perc = ""
   t_iva = ""
   t_total = ""
   Command1.Enabled = False
End Sub
Sub carga()
  q = "select * from emp_02 where [num_mov_int] = " & Val(t_numcomp)
  Set rs = New adodb.Recordset
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
   If rs("tipo_movimiento") = 20 Then
      EXISTE = "S"
      Call pi7
      Command1.Enabled = True
    Else
      MsgBox ("El movimiento no corresponde a un Ingreso de Gastos")
      Call limpia
    End If
  Else
    EXISTE = "N"
    t_fecha = ""
    Call armagrid
  End If
  Set rs = Nothing
End Sub
Sub pi7()
Call armagrid
q = "select * from emp_02 where  [num_mov_int] = " & Val(t_numcomp) & " and [tipo_movimiento] = 20"
Set rs = New adodb.Recordset
rs.Open q, cn1
p = 0
t = 0
While Not rs.EOF
  If p = 0 Then
    t_fecha = rs("fecha")
    t_observaciones = rs("observaciones")
    c_tipocomp.ListIndex = buscaindice(c_tipocomp, rs("id_legajo"))
    p = 1
  End If
  msf1.AddItem rs("observaciones") & Chr$(9) & Format$(rs("IMPORTE"), "######0.00")
  t = t + rs("IMPORTE")
  rs.MoveNext
Wend
Set rs = Nothing

Call sacatotales

End Sub
Private Sub btnacepta_Click()
J = MsgBox("Graba Comprobante", 4)
If J = 6 Then
  Call graba
  
End If
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 3
msf1.ColWidth(0) = 6000
msf1.ColWidth(1) = 2000
msf1.ColWidth(2) = 600
msf1.TextMatrix(0, 0) = "Detalle"
msf1.TextMatrix(0, 1) = "Importe"
msf1.TextMatrix(0, 2) = " "

End Sub



Sub inicia()
Set rs = New adodb.Recordset
q = "select * from g0 "
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
   t_numcomp = rs("ult_num_mov_emp") + 1
Else
  MsgBox ("Error al Inicializar el Formulario")
  Unload Me
End If
c_tipocomp.ListIndex = 0
Call armagrid
Command1.Enabled = False

End Sub

Private Sub c_tipocomp_LostFocus()
 If c_tipocomp.ListIndex < 0 Then
   c_tipocomp.ListIndex = 0
 End If
End Sub



Private Sub Command1_Click()
  Call nivel_acceso(1)
  If para.id_grupo_modulo_actual >= 8 Then
    J = MsgBox("Confirma Borrar movimiento", 4)
    If J = 6 Then
      QUERY = "DELETE FROM emp_02 WHERE [num_mov_int] = " & Val(t_numcomp)
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
    End If
    Call inicia
  Else
    Call sinpermisos
  End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     gen_tools.Show
  
End Select
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call TabEnter2(Me, 5)
End If


End Sub

Private Sub Form_Load()

Call INICIALIZA2(Me)

Call barraesag(Me)

Call carga_empleados(c_tipocomp)
'Load emp_emitemov1
Call inicia

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload emp_emitemov1

End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[INS] Agrega - [ENTER] Modifica - [F5] Elimina - [F7] Imprime - [F9] Termina"
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If
Me.KeyPreview = False

End Sub
Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
 If msf1.Rows > 2 Then
     msf1.RemoveItem (msf1.Row)
     Call sacatotales
 Else
    Call armagrid
 End If
End If


If KeyCode = vbKeyF7 Then
  Call nivel_acceso(1)
  If para.id_grupo_modulo_actual >= 4 Then
    Call imprime
  End If
         
  
End If

If KeyCode = vbKeyInsert Then
   emp_emitegastos1.t_detalle = ""
   emp_emitegastos1.t_cantidad = ""
   
   emp_emitegastos1.Show
   
End If

If KeyCode = vbKeyF9 Then
  Call sacatotales
  t_observaciones.SetFocus
End If


End Sub

Sub graba()
  ' On Error GoTo ERRORGRABA
          tm = 20
          u = "H"
   
   cn1.BeginTrans
  'VERIFICA QUE NO EXISTA EL COMPROBANTE
  Set rs = New adodb.Recordset
  q = "select * from emp_02 where [num_mov_int] = " & Val(t_numcomp)
  rs.Open q, cn1
  If Not rs.BOF And Not rs.EOF Then
    If rs("tipo_movimiento") = 20 Then
               
         QUERY = "delete  emp_02 where [num_mov_int]= " & Val(t_numcomp)
         cn1.Execute QUERY
    Else
       MsgBox ("El movimiento no pertenece a un ingreso de gastos. Imposible modificar.")
       Exit Sub
   End If
  End If
  
  For i = 1 To msf1.Rows - 1
          
          QUERY = "INSERT INTO emp_02([num_mov_int], [id_legajo], [importe], [tipo_movimiento], [fecha], [ubicacion], [observaciones])"
          QUERY = QUERY & " VALUES (" & Val(t_numcomp) & ", " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & ", " & Val(msf1.TextMatrix(i, 1)) & ", " & tm & ", '" & t_fecha & "', '" & u & "', '" & msf1.TextMatrix(i, 0) & ".')"
          'MsgBox (QUERY)
          cn1.Execute QUERY
  Next i
          
QUERY = "update G0 set  [ult_num_mov_emp]=" & Val(t_numcomp)
QUERY = QUERY & " where [sucursal]= 0"
      
cn1.Execute QUERY
cn1.CommitTrans
Set rs = Nothing
 
Call imprime
 
 
Call INICIALIZA2(Me)
Call inicia
t_numcomp.SetFocus

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    emp_emitegastos1.t_detalle = msf1.TextMatrix(msf1.Row, 0)
    emp_emitegastos1.t_cantidad = msf1.TextMatrix(msf1.Row, 1)
    emp_emitegastos1.t_renglon = msf1.Row
    emp_emitegastos1.Show
  End If
End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight
Me.KeyPreview = True

End Sub


Private Sub t_fecha_LostFocus()
If Not IsDate(t_fecha) Then
  t_fecha = Format$(Now, "dd/mm/yyyy")
Else
  t_fecha = Format$(t_fecha, "dd/mm/yyyy")
End If
Call verifica_fechacorte(t_fecha)
End Sub

Private Sub t_numcomp_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)

End Sub

Private Sub t_numcomp_LostFocus()
   t_numcomp = Format$(Val(t_numcomp), "00000000")
   Call carga
End Sub

Private Sub t_observaciones_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   btnacepta.SetFocus
End If
End Sub

Private Sub t_observaciones_LostFocus()
Call NULOS(t_observaciones)
End Sub

Sub sacatotales()
s = 0
For i = 1 To msf1.Rows - 1
   r = Val(msf1.TextMatrix(i, 1))
   s = s + r
 
Next i
t_total = Format$(s, "#####0.00")


End Sub


Private Sub t_total_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If
End Sub

Private Sub t_total_LostFocus()
t_total = Format$(t_total, "######0.00")
End Sub

Sub imprime()
   J = MsgBox("Prepare Impresora y Confirme", 4)
    If J = 6 Then
     Dim c(15) As Double
      c(0) = 0
      c(1) = 1
      For i = 2 To 14
        c(i) = -1
      Next i
     msf1.AddItem ""
     
     Call imprimegrid(msf1, c(), "INGRESO DE GASTOS", "Numero: " & t_numcomp, "Empleado: " & c_tipocomp, "Fecha: " & t_fecha & "   Importe Total: " & t_total, 80, 10, True, False)
     
    End If

End Sub
