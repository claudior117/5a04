VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cgr_verasientos 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MUESTRA ASIENTOS TEMPORALES"
   ClientHeight    =   8775
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   11535
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   11535
      Begin VB.ComboBox c_usuarios 
         Height          =   315
         Left            =   8640
         TabIndex        =   13
         Text            =   "c_usuarios"
         Top             =   720
         Width           =   2775
      End
      Begin VB.ComboBox c_modulo 
         Height          =   315
         ItemData        =   "CGR002.frx":0000
         Left            =   8640
         List            =   "CGR002.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox c_cuenta 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Text            =   "c_cuenta"
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Usuario:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6960
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Modulo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6960
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cuenta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   5
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "CGR002.frx":0085
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "CGR002.frx":0907
         Style           =   1  'Graphical
         TabIndex        =   6
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
      TabIndex        =   4
      Top             =   8520
      Width           =   11985
      _ExtentX        =   21140
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
            TextSave        =   "25/02/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:46"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cgr_verasientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Sub cabecera()
List1.clear
Call cabeceralist(List1)

End Sub

Sub carga()
  Call cabecera
  q = "select * from C_02, C_03 where C_02.[num_interno] = c_03.[num_interno] "
  c = " and "
  If c_cuenta.ListIndex > 0 Then
     q = q & c & " [id_cuenta] = " & c_cuenta.ItemData(c_cuenta.ListIndex)
  End If
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  If c_usuarios.ListIndex > 0 Then
     q = q & c & " [id_usuario] = " & c_usuarios.ItemData(c_usuarios.ListIndex)
  End If
  
  If c_modulo.ListIndex > 0 Then
     q = q & c & " [modulo] = '" & Mid$(c_modulo, 2, 1) & "'"
  End If
  
  q = q & " order by c_02.[num_interno]"
  
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  t = 0
  numint = 0
  While Not rs.EOF
   If numint <> rs("C_02.NUM_INTERNO") Then
      numint = rs("C_02.NUM_INTERNO")
      F = Format$(rs("fecha"), "dd/mm/yyyy")
      ope = rs("c_02.descripcion")
      ni = Format$(rs("c_02.num_interno"), "00000000")
      obs = rs("observaciones")
      List1.AddItem ""
      List1.AddItem "Nro. Interno:..." & ni & "                  Fecha:...." & F
      List1.AddItem "Operacion:......" & ope
      List1.AddItem "Observaciones:.." & obs
      Call armaasiento(numint)
   End If
   rs.MoveNext
  Wend
  
   
End Sub

Sub armaasiento(ByVal i As Long)
q = "select * from c_03, c_01 where [num_interno] = " & i & " and c_03.[id_cuenta] = c_01.[id_cuenta]"
q = q & " order by [ubicacion], [renglon]"
Set rs2 = New ADODB.Recordset
rs2.Open q, cn1
d = Space$(10)
h = Space$(10)
List1.AddItem "========================================================================================="
List1.AddItem "        DETALLE                                  DEBE     HABER   Observaciones"
List1.AddItem "========================================================================================="

While Not rs2.EOF
   ic = rs2("c_03.id_cuenta")
   dc = Format$(Left$(rs2("c_01.descripcion"), 35), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
   RSet d = Format$(rs2("c_03.importe"), "######0.00")
   If rs2("ubicacion") = "D" Then
     e = ""
     e2 = "          "
   Else
     e = "          "
     e2 = ""
   End If
   
   o = Format$(Left$(rs2("c_03.descripcion"), 35), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 List1.AddItem e & "[" & ic & "] " & dc & " " & d & "  " & e2 & o
 rs2.MoveNext
Wend
List1.AddItem "                                           ***********************"
List1.AddItem ""
List1.AddItem ""

End Sub
Private Sub btnacepta_Click()
Call carga
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub










Private Sub c_cuenta_LostFocus()
If c_cuenta.ListIndex < 0 Then
  If Val(c_cuenta) > 0 Then
    c_cuenta.ListIndex = buscaindice(c_cuenta, Val(c_cuenta))
  Else
    c_cuenta.ListIndex = 0
  End If
End If
End Sub

Private Sub c_modulo_LostFocus()
If c_modulo.ListIndex < 0 Then
  c_modulo.ListIndex = 0
End If

End Sub

Private Sub c_usuarios_LostFocus()
If c_usuarios.ListIndex < 0 Then
  c_usuarios.ListIndex = 0
End If

End Sub

Private Sub Form_Load()


Call carga_cuentas_cont(c_cuenta, "C", "D")
c_cuenta.AddItem "<Todos>", 0
c_cuenta.ListIndex = 0

Call carga_usuarios(c_usuarios)
c_usuarios.AddItem "<Todos>", 0
c_usuarios.ListIndex = 0

c_modulo.ListIndex = 0

Call barraesag(Me)


End Sub

Private Sub List1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F8] Borra -  [ENTER] Modifica - [F1] Nuevo - [F7] Imprime"

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF8 Then
 Call nivel_acceso(2)
 If para.id_grupo_modulo_actual >= 8 Then
   
  If Mid$(List1.List(List1.ListIndex), 1, 12) = "Nro. Interno" Then
    J = MsgBox("Confirma Eliminar Asiento " & Mid$(List1.List(List1.ListIndex), 1, 25), 4)
    If J = 6 Then
      nicgr = Val(Mid$(List1.List(List1.ListIndex), 17, 8))
      Set rs = New ADODB.Recordset
      q = "select * from c_02 where [num_interno] = " & nicgr
      rs.MaxRecords = 1
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
        If rs("modulo") <> "G" Then
           nicgr = 0
        End If
      Else
        nicgr = 0
      End If
      
      If nicgr <> 0 Then
        cn1.BeginTrans
        QUERY = "DELETE FROM c_02 WHERE [num_interno] = " & nicgr
        cn1.Execute QUERY
      
        QUERY = "DELETE FROM c_03 WHERE [num_interno] = " & nicgr
        cn1.Execute QUERY
        cn1.CommitTrans
      Else
        MsgBox ("El asiento fue generado por un modulo de Gestion. Borre el movimiento original para eliminarlo")
      End If
      Call carga
    End If
  End If
Else
   Call sinpermisos
End If
End If
If KeyCode = vbKeyF7 Then
  J = MsgBox("Prepare Impresora y Confirme", 4)
  If J = 6 Then
    Call imprimelist(List1, 70, 9, False, True, 5)
  End If

End If


If KeyCode = vbKeyF1 Then
 Call nivel_acceso(2)
 If para.id_grupo_modulo_actual >= 8 Then
    cgr_abmasientos_p.Show
    cgr_abmasientos_p.t_funcion = "A"
  Else
   Call sinpermisos
  End If
End If




If KeyCode = vbKeyF7 Then
  J = MsgBox("Prepare Impresora y Confirme", 4)
  If J = 6 Then
    Call imprimelist(List1, 70, 9, False, True, 5)
  End If

End If


End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call nivel_acceso(2)
 If para.id_grupo_modulo_actual >= 8 Then
   If Mid$(List1.List(List1.ListIndex), 1, 12) = "Nro. Interno" Then
    cgr_abmasientos_p.Show
    cgr_abmasientos_p.t_funcion = "M"
    cgr_abmasientos_p.t_id = Val(Mid$(List1.List(List1.ListIndex), 17, 8))
    cgr_abmasientos_p.LLENACAMPOS
    cgr_abmasientos_p.Show
   End If
  Else
   Call sinpermisos
  End If
End If

End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = ""
  End If
End If
End Sub

Private Sub t_fecha2_GotFocus()
t_fecha2 = ""
End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
  If Not IsDate(t_fecha2) Then
    t_fecha2 = ""
  End If
End If

End Sub
