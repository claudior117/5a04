VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form stk_vercomp 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MOVIMIENTOS EMITIDOS"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordenados por:"
      Height          =   615
      Left            =   240
      TabIndex        =   18
      Top             =   7080
      Width           =   3615
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Detalle"
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   8520
      TabIndex        =   14
      Top             =   1320
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   180813825
      CurrentDate     =   38754
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5055
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8916
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1815
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Width           =   11535
      Begin VB.TextBox t_idp 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   23
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox c_prod 
         Height          =   315
         Left            =   3120
         TabIndex        =   21
         Text            =   "c_prov"
         Top             =   1440
         Width           =   5295
      End
      Begin VB.ComboBox c_vend 
         Height          =   315
         Left            =   1680
         TabIndex        =   16
         Text            =   "c"
         Top             =   840
         Width           =   4815
      End
      Begin VB.ComboBox c_tipocomp 
         Height          =   315
         ItemData        =   "stk008.frx":0000
         Left            =   8640
         List            =   "stk008.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   8640
         MaxLength       =   10
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   8640
         MaxLength       =   10
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Text            =   "c_prov"
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Producto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Obra:"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   9960
         TabIndex        =   15
         Top             =   960
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
         Left            =   6960
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6960
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Proveedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
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
         Picture         =   "stk008.frx":004B
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
         Picture         =   "stk008.frx":08CD
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
      Top             =   8235
      Width           =   11805
      _ExtentX        =   20823
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
            TextSave        =   "26/02/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "04:24 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "stk_vercomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer


Sub carga()
  Call armagrid
  q = "select * from stk_02, stk_03 where stk_02.[num_int] = stk_03.[num_int]  "
  c = " and "
  If c_prov.ListIndex > 0 Then
     q = q & c & " [id_proveedor] = " & c_prov.ItemData(c_prov.ListIndex)
  End If
  
  
  If Val(Mid$(c_tipocomp, 2, 2)) > 0 Then
    q = q & c & " [tipo_comprobante] = " & Val(Mid$(c_tipocomp, 2, 2))
  End If
  
  If t_fecha <> "" And IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If t_fecha2 <> "" And IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
    
   If c_vend.ListIndex > 0 Then
    q = q & c & " [id_obra] = " & c_vend.ItemData(c_vend.ListIndex)
   End If

   Select Case Val(t_idp)
   Case Is = 0
   
   Case Is < 0
     q = q & " and  [descripcion] like '%" & c_prod & "%'"
     
   Case Else
     q = q & " and [id_producto] = " & Val(t_idp)
   End Select
   
   
   If Option1 = True Then
    q = q & " order by [fecha], stk_02.[num_int]"
    Else
    q = q & " order by stk_02.[detalle], [fecha], stk_02.[num_int]"
   End If
 
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  t = 0
  reg = 0
  ni = 0
  While Not rs.EOF
   If ni <> rs("stk_02.num_int") Then
     F = rs("fecha")
     Select Case rs("TIPo_COMProbante")
      Case Is = 1
       CTC = "Ajuste"
         p = " "
         o = " "
      Case Is = 20
       CTC = "Entrada"
         o = " "
         Set rs1 = New ADODB.Recordset
         q = "select * from a1 where [id_proveedor] = " & rs("id_proveedor")
         rs1.Open q, cn1
         If Not rs1.EOF And Not rs1.BOF Then
           p = rs1("denominacion")
         Else
           p = " "
         End If
         Set rs1 = Nothing
         
      Case Is = 30
         CTC = "Salida"
         p = " "
         Set rs1 = New ADODB.Recordset
         q = "select * from a4 where [id_obra] = " & rs("id_obra")
         rs1.Open q, cn1
         If Not rs1.EOF And Not rs1.BOF Then
           o = rs1("descripcion")
         Else
           o = " "
         End If
         Set rs1 = Nothing
     End Select
       
     tc = rs("stk_02.detalle")
     nc = rs("letra") & " " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
     ni = rs("stk_02.num_int")
     
     
     msf1.AddItem F & Chr(9) & CTC & Chr(9) & ni & Chr(9) & tc & Chr(9) & nc & Chr(9) & p & Chr(9) & o
     reg = reg + 1
     Label5 = reg
     Label5.Refresh
   End If
   rs.MoveNext
  Wend
  msf1.AddItem ""
  msf1.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "Comprobantes: " & reg

   
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
msf1.Cols = 7
msf1.ColWidth(0) = 1300
msf1.ColWidth(1) = 700
msf1.ColWidth(2) = 1300
msf1.ColWidth(3) = 3500
msf1.ColWidth(4) = 1700
msf1.ColWidth(5) = 3000
msf1.ColWidth(6) = 3000

msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = "Tipo"
msf1.TextMatrix(0, 2) = "Num. Mov."
msf1.TextMatrix(0, 3) = "Detalle"
msf1.TextMatrix(0, 4) = "Comprobante"
msf1.TextMatrix(0, 5) = "Proveedor"
msf1.TextMatrix(0, 6) = "Obra"
For i = 0 To 6
    msf1.ColAlignment(i) = 1 'izq
Next i

End Sub







Private Sub c_prod_LostFocus()
Select Case c_prod.ListIndex
 Case Is < 0
   t_idp = -1
 Case Is = 0
   t_idp = 0
 Case Else
   t_idp = c_prod.ItemData(c_prod.ListIndex)
End Select
 
  
End Sub

Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
  c_prov.ListIndex = 0
End If
End Sub

Private Sub c_tipocomp_LostFocus()
If c_tipocomp.ListIndex < 0 Then
  c_tipocomp.ListIndex = 0
End If
End Sub

Private Sub c_vend_LostFocus()
If c_vend.ListIndex < 0 Then
  c_vend.ListIndex = 0
End If
End Sub

Private Sub cal1_DblClick()
If cal1.Tag = "1" Then
   t_fecha = cal1
Else
   t_fecha2 = cal1
End If
cal1.Visible = False

End Sub

Private Sub cal1_LostFocus()
If cal1.Tag = "1" Then
   t_fecha = cal1
Else
   t_fecha2 = cal1
End If
cal1.Visible = False

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 4)
  'Case Is = 27
  '      Me.Hide
End Select

End Sub

Private Sub Form_Load()
cal1.Visible = False
Call carga_proveedores(c_prov)
c_prov.AddItem "<Todos>", 0
c_prov.ListIndex = 0

Call carga_obras(c_vend, "A")
c_vend.AddItem "<Todas>", 0
c_vend.ListIndex = 0

'Call carga_productos(c_prod)
c_prod.AddItem "<Todos>", 0
c_prod.ListIndex = 0

Call armagrid
Call barraesag(Me)
Option1 = True
Option2 = True
c_tipocomp.ListIndex = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload vta_clientes
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F1] Cliente -  [F8] Borra - [F3] Cambia Datos  "
Me.KeyPreview = False
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

    For i = 7 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "MOVIMIENTOS STOCK EMITIDOS", "Proveedor:" & c_prov & "    Obra: " & c_vend, "Fecha desde: " & t_fecha & "  Fecha hasta: " & t_fecha2, "Producto: " & c_prod, 50, 8, True, False, "H")
  End If

End If



 If KeyCode = vbKeyF8 Then
  Call nivel_acceso(8)
  If para.id_grupo_modulo_actual >= 8 Then
   J = MsgBox("Confirma Eliminar Comprobante Nro." & msf1.TextMatrix(msf1.RowSel, 2), 4)
   If J = 6 Then
      indice = msf1.RowSel
      
      
      
      MsgBox ("Operacion Terminada")
   End If
  End If
End If


If KeyCode = vbKeyF5 Then
 J = MsgBox("Prepare Impresora y Confirme", 4)
 If J = 6 Then
        Call nivel_acceso(8)
        If para.id_grupo_modulo_actual >= 6 Then
           Set cl_compvta = New comprobantes_venta
           cl_compvta.cargar2 (Val(msf1.TextMatrix(msf1.Row, 8)))
           cl_compvta.imprimir
        End If
  End If
End If



End Sub


Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load stk_cc_detalle
    stk_cc_detalle.t_tipocomp = msf1.TextMatrix(msf1.Row, 1)
    stk_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 2)
    stk_cc_detalle.Show
  End If
End If

End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
Me.KeyPreview = True
msf1.FocusRect = flexFocusLight

End Sub

Private Sub t_fecha_DblClick()
cal1.Visible = True
cal1.Tag = "1"
End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha2_DblClick()
cal1.Visible = True
cal1.Tag = "2"

End Sub

Private Sub t_fecha2_GotFocus()
t_fecha2 = ""
End Sub

Private Sub t_idp_LostFocus()
c_prod.ListIndex = buscaindice(c_prod, Val(t_idp))
End Sub


