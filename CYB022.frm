VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cyb_informeMOV 
   BackColor       =   &H00E0E0E0&
   Caption         =   "INFORME DE MOVIMIENTOS INGRESADOS"
   ClientHeight    =   8760
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordenar por"
      Height          =   615
      Left            =   240
      TabIndex        =   15
      Top             =   7200
      Width           =   3015
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo Movimiento"
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   6600
      TabIndex        =   12
      Top             =   0
      Width           =   5415
      Begin VB.ComboBox c_tipo 
         Height          =   315
         ItemData        =   "CYB022.frx":0000
         Left            =   1080
         List            =   "CYB022.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   600
         Width           =   4215
      End
      Begin VB.ComboBox c_concepto 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Tipo Mov."
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5295
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9340
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   2
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
      TabIndex        =   7
      Top             =   0
      Width           =   6255
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox c_banco 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Banco:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   4
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "CYB022.frx":002D
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "CYB022.frx":08AF
         Style           =   1  'Graphical
         TabIndex        =   5
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
      TabIndex        =   3
      Top             =   8505
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   353
            MinWidth        =   353
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   17639
            MinWidth        =   17639
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
            TextSave        =   "09:39"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cyb_informeMOV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim saldoanterior As Double
Sub carga()
  Call armagrid
  da = 0
  ha = 0
  q = "select * from cyb_04, cyb_06, cyb_01 where  cyb_04.[id_tipomov] = cyb_06.[id_tipomov] and [id_banco] = [id_forma_pago]"
  c = " and "
  
  If c_banco.ListIndex > 0 Then
    q = q & c & " [id_banco] = " & c_banco.ItemData(c_banco.ListIndex)
  End If
  
  If t_fecha <> "" Then
      q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
    
  If t_fecha2 <> "" Then
      q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
    
  
  If c_concepto.ListIndex > 0 Then
     q = q & c & " cyb_04.[id_tipomov] = " & c_concepto.ItemData(c_concepto.ListIndex)
  End If
  
  If c_tipo.ListIndex > 0 Then
    If c_tipo.ListIndex = 1 Then
       q = q & c & " cyb_04.[ubicacion] = 'H'"
    Else
        q = q & c & " cyb_04.[ubicacion] = 'D'"
    End If
  End If
   
  If Option2 = True Then
    q = q & " order by [fecha]"
  Else
    q = q & " order by cyb_04.[id_tipomov]"
  End If
     
  Set rs = New ADODB.Recordset
  'MsgBox (q)
  rs.Open q, cn1
  td = 0
  th = 0
  c = 0
  While Not rs.EOF
     F = rs("fecha")
     ni = Format$(rs("num_mov_banco"), "00000")
     t = rs("cyb_06.abreviatura")
     i = Format$(rs("importe"), "######0.00")
     td = td + Val(i)
     o = rs("detalle")
     b = rs("cyb_01.abreviatura")
     'tm = rs("descripcion")
     u = rs("cyb_04.ubicacion")
     nc = rs("num_comp")
     fd = rs("fecha_dif")
     msf1.AddItem F & Chr(9) & t & Chr(9) & i & Chr(9) & u & Chr(9) & o & Chr(9) & b & Chr(9) & nc & Chr$(9) & rs("entro") & Chr$(9) & fd & Chr$(9) & rs("num_mov_banco")
     c = c + 1
    rs.MoveNext
  Wend
  Set rs = Nothing
  msf1.AddItem "" & Chr(9) & "" & Chr(9) & "==================="
  msf1.AddItem "" & Chr(9) & "" & Chr(9) & Format$(td, "######0.00")
  msf1.AddItem ""
  msf1.AddItem "" & Chr(9) & "Cantidad Registros: " & Chr(9) & c
  
End Sub

Private Sub btnacepta_Click()

  Call carga
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 10
msf1.ColWidth(0) = 1100
msf1.ColWidth(1) = 1000
msf1.ColWidth(2) = 1200
msf1.ColWidth(3) = 600
msf1.ColWidth(4) = 2000
msf1.ColWidth(5) = 2000
msf1.ColWidth(6) = 1000
msf1.ColWidth(7) = 500
msf1.ColWidth(8) = 1000
msf1.ColWidth(9) = 1000

msf1.TextMatrix(0, 0) = "Fecha Dif. "
msf1.TextMatrix(0, 1) = "Tipo "
msf1.TextMatrix(0, 2) = "Importe"
msf1.TextMatrix(0, 3) = "D/H"
msf1.TextMatrix(0, 4) = "Detalle"
msf1.TextMatrix(0, 5) = "Banco"
msf1.TextMatrix(0, 6) = "Num. Comp."
msf1.TextMatrix(0, 7) = "Entro"
msf1.TextMatrix(0, 8) = "Fecha dif."
msf1.TextMatrix(0, 9) = "Nro. Int"
End Sub




Private Sub c_concepto_LostFocus()
If c_concepto.ListIndex < 0 Then
  c_concepto.ListIndex = 0
End If
End Sub




Private Sub c_tipo_LostFocus()
If c_tipo.ListIndex < 0 Then
  c_tipo.ListIndex = 0
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub


Private Sub Form_Load()
Call carga_formas_pago(c_banco, "B")
c_banco.AddItem "<Todos>", 0
c_banco.ListIndex = 0
Call carga_mov_banco(c_concepto)
c_concepto.AddItem "<Todos>", 0
c_concepto.ListIndex = 0
c_tipo.ListIndex = 0
Option2 = True
Call armagrid
Call barraesag(Me)
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime "
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
    c(6) = 6
    c(7) = 7
    c(8) = 8
      
    For i = 9 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "INFORME DE MOVIMIENTOS BANCARIOS", "Tipo:" & c_concepto & " / " & c_tipo, "Periodo...: " & t_fecha & " " & t_fecha2, "Banco.....:" & c_banco, 76, 8, True, False)
  End If
    
End If


If KeyCode = vbKeyF8 Then
 'borrar mov.
 J = MsgBox("Confirma Borrar Movimiento", 4)
 If J = 6 Then
    ni = Val(msf1.TextMatrix(msf1.Row, 10))
    Set cl_banco = New bancos
    cl_banco.borrar (ni)
    Set cl_banco = Nothing
 End If
    
End If


If KeyCode = vbKeyF5 Then
    cyb_concilia.t_id = Val(msf1.TextMatrix(msf1.Row, 10))
    cyb_concilia.t_mov = msf1.TextMatrix(msf1.Row, 3)
    cyb_concilia.t_entro = msf1.TextMatrix(msf1.Row, 7)
    If cyb_concilia.t_entro = "S" Then
      cyb_concilia.t_fecha = msf1.TextMatrix(msf1.Row, 8)
    Else
      cyb_concilia.t_fecha = msf1.TextMatrix(msf1.Row, 0)
    End If
    cyb_concilia.Show
End If


If KeyCode = vbKeyF3 Then
    q = "select * from cyb_04 where [num_mov_banco] = " & Val(msf1.TextMatrix(msf1.Row, 10))
    Set rs = New ADODB.Recordset
    rs.Open q, cn1
    If Not rs.EOF And Not rs.BOF Then
       If rs("modulo") = "C" And rs("num_mov_int") > 0 Then
          q = "select * from a5 where [num_int] = " & rs("num_mov_int")
          Set rs1 = New ADODB.Recordset
          rs1.Open q, cn1
          If Not rs1.EOF And Not rs1.BOF Then
            If rs1("id_tipocomp") = 50 Then
               Load cc_detalle
               cc_detalle.t_numint = rs1("num_int")
               cc_detalle.Show
            End If
          End If
          Set rs1 = Nothing
       End If
     End If
    Set rs = Nothing
End If


End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  cyb_cc_detalleb.t_numint = msf1.TextMatrix(msf1.Row, 9)
  cyb_cc_detalleb.Show
End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight

End Sub

Private Sub t_fecha_DblClick()
cal1.Visible = True
End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call carga
End If
End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = ""
  End If
End If
End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
  If Not IsDate(t_fecha2) Then
    t_fecha2 = ""
  End If
End If

End Sub
