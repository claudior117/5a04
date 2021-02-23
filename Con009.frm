VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form con_retperc 
   BackColor       =   &H00E0E0E0&
   Caption         =   "INFORME DE RETENCIONES Y PERCEPCIONES DE COMPRA"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   6240
      TabIndex        =   11
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox c_tipo 
         Height          =   315
         ItemData        =   "Con009.frx":0000
         Left            =   1680
         List            =   "Con009.frx":000D
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         Caption         =   "Tipo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   2880
      TabIndex        =   9
      Top             =   600
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   111738881
      CurrentDate     =   38750
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3615
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   3
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "Con009.frx":0035
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "Con009.frx":08B7
         Style           =   1  'Graphical
         TabIndex        =   4
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
      TabIndex        =   2
      Top             =   8235
      Width           =   11880
      _ExtentX        =   20955
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
            TextSave        =   "23/02/2021"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:28 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5535
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9763
      _Version        =   393216
   End
End
Attribute VB_Name = "con_retperc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984




Private Sub btnacepta_Click()
Call armagrid
Select Case c_tipo.ListIndex
Case Is = 0
   Call ret
   msf1.AddItem ""
   msf1.AddItem ""
   Call perc
Case Is = 1
   Call ret
Case Is = 2
   Call perc
End Select

End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Sub ret()
msf1.AddItem "" & Chr$(9) & "************ RETENCIONES ************"
msf1.AddItem ""

Set rs1 = New adodb.Recordset
q = "select * from g2 where [id_tipo_comp] >= 95 and [id_tipo_comp] < 100"
rs1.Open q, cn1
tr = 0
While Not rs1.EOF
 p = 0
 q = "select * from a5, g2, a1, g3 where  [id_tipocomp] = [id_tipo_comp] and a5.[id_proveedor] = a1.[id_proveedor] and a1.[cod_tipoiva] = g3.[cod_tipoiva] AND [Id_tipocomp] = " & rs1("id_tipo_comp")
 c = " and "
 If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
 End If
 If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
 End If
 q = q & " order by [fecha]"
 Set rs = New adodb.Recordset
 rs.Open q, cn1
 tt = 0
 While Not rs.EOF
    If p = 0 Then
      msf1.AddItem "" & Chr(9) & rs1("descripcion")
      p = 1
    End If
     F = Format$(rs("fecha"), "dd/mm/yy")
     tc = rs("g2.abreviatura")
     nc = rs("letra") & " " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
     If rs("moneda") = "P" Then
        c5 = 1
     Else
        c5 = rs("cotiz_dolar")
     End If
     
     t = Format$(rs("total") * c5, "######0.00")
     
     tt = tt + Val(t)
     
     msf1.AddItem F & Chr(9) & rs("proveedor05") & Chr(9) & rs("cuit05") & " " & rs("g3.abreviatura") & Chr(9) & tc & " " & nc & Chr(9) & t & Chr(9) & Format$(rs("num_int"), "00000")
     rs.MoveNext
  Wend
  If p > 0 Then
    msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "" & Chr(9) & "______________________"
    msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "Totales:" & Chr(9) & Format$(tt, "######0.00")
    tr = tr + tt
  End If
  rs1.MoveNext
Wend
msf1.AddItem " "
msf1.AddItem " " & Chr(9) & " TOTALE DE RETENCIONES ------------>" & Chr(9) & " " & Chr(9) & "" & Chr(9) & Format$(tr, "######0.00")

End Sub

Sub perc()
msf1.AddItem "" & Chr$(9) & "************ PERCEPCIONES ************"
msf1.AddItem ""

Set rs1 = New adodb.Recordset
q = "select * from A12 where [tipo12] = 'P' and [impuesto12] = 'I'"
rs1.Open q, cn1
tr = 0
While Not rs1.EOF
 p = 0
 q = "select * from a5, g2, a1, g3, a13, a12 where  [GRABADO] <> 'N' AND [id_tipocomp] = [id_tipo_comp] and a5.[id_proveedor] = a1.[id_proveedor] and a1.[cod_tipoiva] = g3.[cod_tipoiva] AND a5.[num_int] = a13.[num_int] and A13.[id_percepcion] = " & rs1("id_percepcion")
 c = " and "
 If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
 End If
 If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
 End If
 q = q & " order by [fecha]"
 Set rs = New adodb.Recordset
 rs.Open q, cn1
 tt = 0
 While Not rs.EOF
    If p = 0 Then
      msf1.AddItem "" & Chr(9) & rs1("descripcion")
      p = 1
    End If
     F = Format$(rs("fecha"), "dd/mm/yy")
     tc = rs("g2.abreviatura")
     nc = rs("letra") & " " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
     If rs("moneda") = "P" Then
        c5 = 1
     Else
        c5 = rs("cotiz_dolar")
     End If
     
     If rs("grabado") = "S" Then
       t = Format$(rs("importe") * c5, "######0.00")
     Else
        t = Format$(-rs("importe") * c5, "######0.00")
     End If
     tt = tt + Val(t)
     
     msf1.AddItem F & Chr(9) & rs("proveedor05") & Chr(9) & rs("cuit05") & " " & rs("g3.abreviatura") & Chr(9) & tc & " " & nc & Chr(9) & t & Chr(9) & Format$(rs("A5.num_int"), "00000")
     rs.MoveNext
  Wend
  If p > 0 Then
    msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "" & Chr(9) & "______________________"
    msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "Totales:" & Chr(9) & Format$(tt, "######0.00")
    tr = tr + tt
  End If
  rs1.MoveNext
Wend
msf1.AddItem " "
msf1.AddItem " " & Chr(9) & " TOTAL DE PERCEPCIONES ------------>" & Chr(9) & " " & Chr(9) & "" & Chr(9) & Format$(tr, "######0.00")


End Sub



Private Sub c_tipo_LostFocus()
If c_tipo.ListIndex < 0 Then
  c_tipo.ListIndex = 0
End If

End Sub

Private Sub cal1_DblClick()
If cal1.Tag = "1" Then
  t_fecha = cal1.Value
Else
  t_fecha2 = cal1.Value
End If
cal1.Visible = False
End Sub

Private Sub cal1_LostFocus()
cal1.Visible = False
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     gen_tools.Show
End Select
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 6
msf1.ColWidth(0) = 900
msf1.ColWidth(1) = 5000
msf1.ColWidth(2) = 1800
msf1.ColWidth(3) = 2200
msf1.ColWidth(4) = 1100
msf1.ColWidth(5) = 1100


msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = "Proveedor"
msf1.TextMatrix(0, 2) = "Cuit "
msf1.TextMatrix(0, 3) = "Tipo y Nro.Comprob."
msf1.TextMatrix(0, 4) = "Importe  "
msf1.TextMatrix(0, 5) = "Num.Int."

For i = 0 To 3
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 4 To 5
  msf1.ColAlignment(i) = 9 'der
Next i

End Sub

Private Sub Form_Load()

Call barraesag(Me)
cal1.Visible = False
Call armagrid
c_tipo.ListIndex = 0
End Sub



Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - [F6] Archivo Texto - [F11] Excel"

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
    
    For i = 5 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "LISTADO DE RETENCIONES y PERCEPCIONES por COMPRAS", "", "Periodo: " & t_fecha & " : " & t_fecha2, "", 95, 6, True, False)
  End If

End If


If KeyCode = vbKeyF6 Then
  Dim c2(15) As Double
    c(0) = 0
    c(1) = 1
    c(2) = 2
    c(3) = 3
    c(4) = 4
    
    For i = 5 To 14
      c(i) = -1
    Next i
    Call exportagrid(msf1, c(), "LISTADO DE RETENCIONES y PERCEPCIONES por COMPRAS", "", "Periodo: " & t_fecha & " : " & t_fecha2, "", True, False, para.archivo_exportacion)

End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load cc_detalle
    cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 5)
    cc_detalle.Show
  End If
End If

End Sub

Private Sub t_fecha_DblClick()
cal1.Visible = True
cal1.Tag = "1"


End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = Format$(Now, "dd/mm/yyyy")
  End If
End If
End Sub

Private Sub t_fecha2_DblClick()
cal1.Visible = True
cal1.Tag = "2"

End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
  If Not IsDate(t_fecha2) Then
    t_fecha2 = Format$(Now, "dd/mm/yyyy")
  End If
End If

End Sub
