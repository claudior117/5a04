VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_retyperc 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORME DE RETENCIONES y PERCEPCIONES RECIBIDAS"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   12180
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   7080
      TabIndex        =   11
      Top             =   0
      Width           =   4695
      Begin VB.ComboBox c_imp 
         Height          =   315
         Left            =   1440
         TabIndex        =   15
         Top             =   720
         Width           =   3135
      End
      Begin VB.ComboBox c_comp 
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Impuesto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Tipo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   3720
      TabIndex        =   9
      Top             =   120
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   169410561
      CurrentDate     =   38750
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3255
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
         Picture         =   "VTA016.frx":0000
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
         Picture         =   "VTA016.frx":0882
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
      Top             =   8640
      Width           =   12180
      _ExtentX        =   21484
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
            TextSave        =   "12/04/2021"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:40 a.m."
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
Attribute VB_Name = "vta_retyperc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim totperc As Double
Dim totret As Double
Const l = "-------------------------------------------"

Sub carga()
Call armagrid
ttp = 0
ttr = 0
If c_imp.ListIndex > 0 Then
  msf1.AddItem "" & Chr$(9) & "IMPUESTO-------->" & Chr$(9) & UCase(c_imp)
  ttp = sacaperc(c_imp.ListIndex)
  msf1.AddItem ""
  ttr = SACAret(c_imp.ListIndex)
  If ttp + ttr > 0 Then
    msf1.AddItem ""
    msf1.AddItem "" & Chr$(9) & Chr$(9) & "TOTAL RET.+PERC.---->" & Chr$(9) & UCase(c_imp) & Chr$(9) & Chr$(9) & Chr$(9) & Format$(ttp + ttr, "######0.00")
    msf1.AddItem ""
  End If
Else
  For i = 1 To 4
    tp = 0
    tr = 0
    msf1.AddItem "" & Chr$(9) & "IMPUESTO-------->" & Chr$(9) & UCase(c_imp.List(i))
    tp = sacaperc(i)
    ttp = ttp + tp
    msf1.AddItem ""
    tr = SACAret(i)
    ttr = ttr + tr
    If tp + tr > 0 Then
      msf1.AddItem ""
      msf1.AddItem "" & Chr$(9) & Chr$(9) & "TOTAL RET.+PERC.------>" & Chr$(9) & UCase(c_imp.List(i)) & Chr$(9) & Chr$(9) & Chr$(9) & Format$(tp + tr, "######0.00")
      msf1.AddItem ""
    End If
  Next i
End If
  
  
   
   
   
End Sub
Function sacaperc(ByVal LI As Integer) As Double
  
tp = 0
If c_comp.ListIndex = 0 Or c_comp.ListIndex = 1 Then 'percePCIONES
  msf1.AddItem "" & Chr$(9) & l & Chr$(9) & l & Chr$(9) & "PERCEPCIONES"
    Select Case LI
   Case Is = 1
     'iva
     tp = buscaperc("I")
    
   Case Is = 2
     'ib
     tp = buscaperc("B")
     
   Case Is = 3
     'gan
     tp = buscaperc("G")
     
   Case Is = 4
        'suss
     tp = buscaperc("S")
     
   End Select
 End If
 sacaperc = tp
End Function
Function SACAret(ByVal LI As Integer) As Double
 ttr = 0
 If c_comp.ListIndex = 0 Or c_comp.ListIndex = 2 Then 'retenciones
  msf1.AddItem "" & Chr$(9) & l & Chr$(9) & l & Chr$(9) & "RETENCIONES"
  
  Select Case LI
   Case Is = 1
     'iva
     ttr = buscaRET2("I")
     
   Case Is = 2
     'ib
     ttr = buscaRET2("B")
   Case Is = 3
     'gan
     ttr = buscaRET2("G")
     ttr = totperc
   Case Is = 4
        'suss
     ttr = buscaRET2("S")
     
   End Select
 End If
SACAret = ttr
End Function
Function buscaperc(ByVal t As String) As Double

Set rs1 = New ADODB.Recordset
q = "select * from a12 where [tipo12] = 'P' and [impuesto12] = '" & t & "'"
rs1.Open q, cn1
totperc = 0
Select Case t
Case Is = "I"
   dr = "IVA"
Case Is = "B"
   dr = "IB"
Case Is = "G"
   dr = "GAN"
Case Is = "S"
   dr = "SEG.SOC."
End Select
  
tp = 0
While Not rs1.EOF
 
 p = 0
 q = "select * from a5, g2, g3, a1, a13  where  [GRABADO] <> 'N' AND [id_tipocomp] = [id_tipo_comp] and a5.[id_proveedor] = a1.[id_proveedor]  AND a5.[num_int] = a13.[num_int] and a13.[id_percepcion] = " & rs1("id_percepcion") & " and a1.[cod_tipoiva] = g3.[cod_tipoiva]"
 c = " and "
 If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
 End If
 If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
 End If
 q = q & " order by [fecha]"
 'MsgBox (q)
 Set rs = New ADODB.Recordset
 rs.Open q, cn1
 tt = 0
 While Not rs.EOF
    If p = 0 Then
      msf1.AddItem "" & Chr(9) & UCase(rs1("descripcion"))
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
     
     msf1.AddItem F & Chr(9) & "" & Chr$(9) & rs("denominacion") & Chr(9) & rs("cuit") & " " & rs("G3.abreviatura") & Chr(9) & tc & " " & nc & Chr(9) & "" & Chr(9) & t & Chr(9) & Format$(rs("A5.num_int"), "00000") & Chr$(9) & "C"
     rs.MoveNext
  Wend
    If tt > 0 Then
     msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "" & Chr(9) & "" & Chr(9) & "______________________"
     msf1.AddItem "" & Chr(9) & "" & Chr(9) & "Total " & rs1("descripcion") & Chr(9) & " " & Chr$(9) & Chr(9) & Chr$(9) & Format$(tt, "######0.00")
     msf1.AddItem ""
  End If
  tp = tp + tt
  Set rs = Nothing
  rs1.MoveNext
 Wend
 If tp > 0 Then
   msf1.AddItem " " & Chr(9) & Chr$(9) & " TOTAL  PERCEPCIONES " & dr & Chr(9) & l & Chr(9) & l & Chr(9) & l & Chr$(9) & Format$(tp, "######0.00")
 End If
 buscaperc = tp
End Function
  
  
  
Function buscaRET2(ByVal t As String) As Double

Select Case t
Case Is = "I"
   dr = "IVA"
   cc = 101
Case Is = "B"
   dr = "IB"
   cc = 100
Case Is = "G"
   dr = "GAN"
   cc = 102
Case Is = "S"
   dr = "SEG.SOC."
   cc = 103
End Select


q = "select * from VTA_02, vta_06, VTA_01 where  vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and VTA_02.[id_CLIENTE] = VTA_01.[id_CLIENTE] and vta_02.[sucursal] = vta_06.[sucursal] "
c = " and "
  
If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
End If
  
If IsDate(t_fecha2) Then
    q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
End If
q = q & " and (vta_02.[id_tipocomp] = " & cc & " or (vta_02.[id_tipocomp] >= 205 and vta_02.[id_tipocomp] <= 207)" & " or (vta_02.[id_tipocomp] = 400))"
q = q & " order by [fecha]"
  
Set rs = New ADODB.Recordset
rs.Open q, cn1
tt = 0
p = 0
While Not rs.EOF
    If p = 0 Then
      msf1.AddItem "" & Chr(9) & UCase("RETENCIONES " & dr)
      p = 1
    End If
    F = Format$(rs("fecha"), "dd/mm/yy")
     tc = rs("abreviatura")
     nc = rs("letra") & " " & Format$(rs("vta_02.sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
     If rs("vta_02.moneda") = "P" Then
       c5 = 1
     Else
       c5 = rs("cotizacion_dolar")
     End If
     c = rs("cuit")
     If rs("vta_02.id_tipocomp") = cc Then
       i = Format$(rs("total") * c5, "######0.00")
     Else
       
            'busco retencions
            q = "select * from vta_012, a12 where [id_retencion] = [id_percepcion] and [num_int] = " & rs("num_int") & " and [impuesto12] = '" & t & "'"
            Set rs1 = New ADODB.Recordset
            rs1.Open q, cn1
            ret = 0
            While Not rs1.EOF
                ret = ret + rs1("importe")
                rs1.MoveNext
            Wend
            Set rs1 = Nothing
            
            If rs("vta_02.id_tipocomp") = 207 Then
              ret = -ret
            End If

            i = Format$(ret * c5, "######0.00")
     End If
     tt = tt + Val(i)
     If Val(i) <> 0 Then
       msf1.AddItem F & Chr(9) & "" & Chr$(9) & rs("denominacion") & Chr(9) & c & Chr$(9) & tc & " " & nc & Chr(9) & s & Chr(9) & i & Chr(9) & Format$(rs("num_int"), "00000") & Chr$(9) & "V"
     End If
     rs.MoveNext
  Wend
  Set rs = Nothing
  If tt > 0 Then
     msf1.AddItem " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "" & Chr(9) & "" & Chr(9) & "______________________"
     msf1.AddItem "" & Chr(9) & "" & Chr(9) & "Total retenciones " & dr & Chr(9) & " " & Chr$(9) & Chr(9) & Chr$(9) & Format$(tt, "######0.00")
     msf1.AddItem ""
  End If
  Set rs = Nothing
  buscaRET2 = tt
End Function

Private Sub btnacepta_Click()
espere.Show
espere.Refresh
Call carga
Unload espere

End Sub

Private Sub btnsale_Click()
Unload Me
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
     Unload Me
End Select
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 9
msf1.ColWidth(0) = 900
msf1.ColWidth(1) = 1700
msf1.ColWidth(2) = 2700
msf1.ColWidth(3) = 1800
msf1.ColWidth(4) = 2200
msf1.ColWidth(5) = 1100
msf1.ColWidth(6) = 1100
msf1.ColWidth(7) = 1100
msf1.ColWidth(8) = 700



msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = "Tipo Impuesto"
msf1.TextMatrix(0, 2) = "Cliente/Proveedor"
msf1.TextMatrix(0, 3) = "Cuit"
msf1.TextMatrix(0, 4) = "Tipo y Nro.Comprob."
msf1.TextMatrix(0, 5) = "Imponible"
msf1.TextMatrix(0, 6) = "Impuesto"
msf1.TextMatrix(0, 7) = "Num.Int."
msf1.TextMatrix(0, 8) = "Modulo "

For i = 0 To 4
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 5 To 7
  msf1.ColAlignment(i) = 9 'der
Next i

End Sub

Private Sub Form_Load()
Call barraesag(Me)
cal1.Visible = False
Call armagrid
Call cargaret

End Sub
Sub cargaret()
'impuestos
c_imp.clear
c_imp.AddItem "<Todos>", 0
c_imp.AddItem "Iva", 1
c_imp.AddItem "Ing.Brutos", 2
c_imp.AddItem "Ganancias", 3
c_imp.AddItem "Seg. Social", 4
c_imp.ListIndex = 0

c_comp.clear
c_comp.AddItem "<Todos>", 0
c_comp.AddItem "Percepciones", 1
c_comp.AddItem "Retenciones", 2
c_comp.ListIndex = 0
End Sub


Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - "

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 8
    c(1) = 0
    c(2) = 1
    c(3) = 2
    c(4) = 3
    c(5) = 4
    c(6) = 5
    c(7) = 6
    c(8) = 7
    For i = 9 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "LISTADO DE RETENCIONES y PERCEPCIONES RECIBIDAS por VENTAS", "", "Periodo: " & t_fecha & " : " & t_fecha2, "", 95, 6, True, False)
  End If

End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    If msf1.TextMatrix(msf1.Row, 8) = "V" Then
     Load vta_cc_detalle
     vta_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 7)
     vta_cc_detalle.Show
   Else
     Load cc_detalle
     cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 7)
     cc_detalle.Show
   
   End If
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
