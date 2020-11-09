VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form prod_vercomp 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADMINISTRADOR DE COMPROBANTES INGRESADOS"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12075
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5295
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9340
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1695
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Width           =   11535
      Begin VB.ComboBox c_usuario 
         Height          =   315
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox t_producto 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   14
         Top             =   1200
         Width           =   3015
      End
      Begin VB.ComboBox c_tipocomp 
         Height          =   315
         ItemData        =   "Pro002.frx":0000
         Left            =   8640
         List            =   "Pro002.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox c_obra 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Text            =   "c_obra"
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Usuario:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6960
         TabIndex        =   17
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Producto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Tipo Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6960
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Obra:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   6
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "Pro002.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "Pro002.frx":0886
         Style           =   1  'Graphical
         TabIndex        =   7
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
      TabIndex        =   5
      Top             =   8385
      Width           =   12075
      _ExtentX        =   21299
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
            TextSave        =   "05/05/2014"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "16:44"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "prod_vercomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim indice As Long


Sub carga()
  Call armagrid
  q = "select * from pro_01, pro_03, a4, g1 where pro_01.[id_tipocomp] = pro_03.[id_tipocomp] and pro_01.[id_obra] = a4.[id_obra] and pro_01.[id_usuario] = g1.[id_usuario] "
  c = " and "
  If C_OBRA.ListIndex > 0 Then
     q = q & c & " pro_01.[id_obra] = " & C_OBRA.ItemData(C_OBRA.ListIndex)
  End If
  
  If c_tipocomp.ListIndex > 0 Then
   
     tc = c_tipocomp.ItemData(c_tipocomp.ListIndex)
   
    q = q & c & " pro_01.[id_tipocomp] = " & tc
  End If
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
    
  If c_usuario.ListIndex > 0 Then
    q = q & c & " pro_01.[id_usuario] = " & c_usuario.ItemData(c_usuario.ListIndex)
  End If
  
    
  q = q & " order by [fecha]"
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
     espere!ProgressBar1.Max = 100
     espere!ProgressBar1.Min = 1
     espere.Show
     espere.Refresh
  End If
  t = 0
  pb = 1
  While Not rs.EOF
     espere!ProgressBar1 = pb
     
     F = rs("fecha")
     tc = rs("abreviatura")
     nc = tc & " " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
     cp = Format$(rs("pro_01.id_obra"), "0000")
     p = rs("a4.descripcion")
     ni = rs("num_int")
     u = rs("usuario")
     o = rs("observaciones")
     msf1.AddItem F & Chr(9) & cp & Chr(9) & p & Chr(9) & nc & Chr(9) & o & Chr(9) & u & Chr(9) & rs("num_int")
    rs.MoveNext
    pb = pb + 1
    If pb > 100 Then
      pb = 1
    End If
  Wend
  espere.Hide
   
End Sub
Sub carga2()
  Call armagrid
  q = "select * from pro_01, pro_02, pro_03, a4, g1 where pro_01.[id_tipocomp] = pro_03.[id_tipocomp] and  pro_01.[num_int] = pro_02.[num_int] and  pro_01.[id_obra] = a4.[id_obra] and pro_01.[id_usuario] = g1.[id_usuario] "
  c = " and "
  If C_OBRA.ListIndex > 0 Then
     q = q & c & " pro_01.[id_obra] = " & C_OBRA.ItemData(C_OBRA.ListIndex)
  End If
  
  If c_tipocomp.ListIndex > 0 Then
    q = q & c & " pro_01.[id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex)
  End If
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
    
  If c_usuario.ListIndex > 0 Then
    q = q & c & " [id_usuario] = " & c_usuario.ItemData(c_usuario.ListIndex)
  End If
   
  If t_producto <> "" Then
    q = q & c & "[pro_02.descripcion] like '%" & t_producto & "%'"
  End If
  
  q = q & " order by [fecha]"
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
     espere!ProgressBar1.Max = 100
     espere!ProgressBar1.Min = 1
     espere.Show
     espere.Refresh
  End If
  t = 0
  pb = 1
  While Not rs.EOF
     espere!ProgressBar1 = pb
     
     F = rs("fecha")
     tc = rs("abreviatura")
     nc = tc & " " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
     cp = Format$(rs("pro_01.id_obra"), "0000")
     p = rs("a4.descripcion")
     ni = rs("pro_01.num_int")
     u = rs("usuario")
     o = rs("pro_01.observaciones")
     msf1.AddItem F & Chr(9) & cp & Chr(9) & p & Chr(9) & nc & Chr(9) & o & Chr(9) & u & Chr(9) & rs("pro_01.num_int")
    rs.MoveNext
    pb = pb + 1
    If pb > 100 Then
      pb = 1
    End If
  Wend
  
  
  espere.Hide
End Sub
Private Sub btnacepta_Click()
If t_producto <> "" Then
    Call carga2
Else
    Call carga
End If
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 7
msf1.ColWidth(0) = 1100
msf1.ColWidth(1) = 600 'cod obra
msf1.ColWidth(2) = 3000
msf1.ColWidth(3) = 2000
msf1.ColWidth(4) = 2500
msf1.ColWidth(5) = 1000
msf1.ColWidth(6) = 1000

msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = ""
msf1.TextMatrix(0, 2) = "Obra"
msf1.TextMatrix(0, 3) = "Nro. Comprobante"
msf1.TextMatrix(0, 4) = "Observaciones"
msf1.TextMatrix(0, 5) = "Usuario"
msf1.TextMatrix(0, 6) = "Num.Int."

For i = 0 To 6
  msf1.ColAlignment(i) = 1 'izq
Next i

End Sub









Private Sub c_obra_LostFocus()
If C_OBRA.ListIndex < 0 Then
 C_OBRA.ListIndex = 0
End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub

Private Sub Form_Load()
Load espere
Call carga_obras(C_OBRA, "A")
C_OBRA.AddItem "<Todas>", 0
C_OBRA.ListIndex = 0

Call carga_tipocompprod(c_tipocomp)
c_tipocomp.AddItem "<Todos>", 0
c_tipocomp.ListIndex = 0

Call carga_usuarios(c_usuario)
c_usuario.AddItem "<Todas>", 0
c_usuario.ListIndex = 0

Call armagrid
Call barraesag(Me)


End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload espere
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime - [ENTER] Detalla - [F8] Borra - [F5] Imprime Comp. - [F3] Cambia Pago"
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 0
    c(1) = 1
    c(2) = 2
    c(3) = 3
    c(4) = 4
    c(5) = 5

    For i = 6 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "COMPROBANTES EMITIDOS PRODUCCION", "", "", "", 72, 8, True, False)
  End If

End If


If KeyCode = vbKeyF8 Then
 Call nivel_acceso(6)
 If para.id_grupo_modulo_actual >= 8 Then

   J = MsgBox("Confirma Eliminar comprobante " & msf1.TextMatrix(msf1.RowSel, 2), 4)
   If J = 6 Then
      indice = msf1.RowSel
      Set cl_compprod = New comprobantes_produccion
      cl_compprod.cargar2 (Val(msf1.TextMatrix(indice, 6)))
      cl_compprod.borrocomp
      Set cl_compprod = Nothing
      MsgBox ("Operacion Terminada")
   End If
  End If
End If


If KeyCode = vbKeyF5 Then
 J = MsgBox("Prepare Impresora y Confirme", 4)
 If J = 6 Then
        Call nivel_acceso(6)
        If para.id_grupo_modulo_actual >= 6 Then
           Set cl_compprod = New comprobantes_produccion
           cl_compprod.cargar2 (Val(msf1.TextMatrix(msf1.Row, 6)))
           cl_compprod.imprimir
        End If
  End If
End If


End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load prod_cc_detalle
    prod_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 6)
    prod_cc_detalle.Show
  End If
End If

End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight

End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha2_GotFocus()
t_fecha2 = ""
End Sub

Private Sub t_producto_GotFocus()
t_producto = ""
End Sub
