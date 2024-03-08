VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_perc 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PERCEPCIONES POR VENTAS"
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
      Top             =   120
      Width           =   4695
      Begin VB.ComboBox c_concepto 
         Height          =   315
         Left            =   1440
         TabIndex        =   14
         Top             =   720
         Width           =   3135
      End
      Begin VB.ComboBox c_imp 
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Concepto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Impuesto:"
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
      StartOfWeek     =   172818433
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
         Picture         =   "VTA064.frx":0000
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
         Picture         =   "VTA064.frx":0882
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
            TextSave        =   "08/03/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "12:38 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5415
      Left            =   0
      TabIndex        =   10
      Top             =   1680
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9551
      _Version        =   393216
   End
End
Attribute VB_Name = "vta_perc"
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

q = "select * from a12 where tipo12 = 'P'"
c = " and "
If c_imp.ListIndex > 0 Then
  Select Case c_imp.ListIndex
   Case Is = 1
     'iva
     tp = "I"
     dr = "IVA "
   Case Is = 2
     'ib
     tp = "B"
      dr = "INGRESOS BRUTOS "
   Case Is = 3
     'gan
     tp = "G"
      dr = "GANANCIAS "
     
   Case Is = 4
        'suss
     tp = "S"
      dr = "SUSS "
     
   Case Is = 5
        'otros
     tp = "O"
      dr = "OTROS "
   End Select
   
   q = q & " and impuesto12 = '" & tp & "'"
   
 End If

 If c_concepto.ListIndex > 0 Then
   q = q & " and id_percepcion = " & c_concepto.ItemData(c_concepto.ListIndex)
 End If

 q = q & " order by [impuesto12], [id_percepcion]"
 
Set rs = New ADODB.Recordset
rs.Open q, cn1
tt = 0
ti = 0
tcp = 0
timp = ""
codper = 0
v = 0
While Not rs.EOF
    
    If codper <> rs("id_percepcion") Or v = 0 Then
     msf1.AddItem ""
     msf1.AddItem "     " & Chr(9) & rs("descripcion")
     codper = rs("id_percepcion")
     v = 1
    End If
 
    q = "select * from vta_02, vta_016 where vta_02.num_int = vta_016.num_int and vta_016.id_percepcion = " & rs("id_percepcion")
    
    If IsDate(t_fecha) Then
     q = q & " and datevalue([fecha]) >= datevalue('" & t_fecha & "')"
    End If
 
    If IsDate(t_fecha2) Then
      q = q & " and  datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
    End If
 
    
    
    Set rs2 = New ADODB.Recordset
    rs2.Open q, cn1
    tcp = 0
    t = "0"
    While Not rs2.EOF
        F = Format$(rs2("fecha"), "dd/mm/yy")
        nc = rs2("letra") & " " & Format$(rs2("sucursal"), "0000") & "-" & Format$(rs2("num_comp"), "00000000")
        If rs2("moneda") = "P" Then
           c5 = 1
        Else
           c5 = rs2("cotiz_dolar")
        End If
        
        If rs2("grabado") = "S" Then
          t = Format$(rs2("importe") * c5, "######0.00")
        Else
           t = Format$(-rs2("importe") * c5, "######0.00")
        End If
        tcodperc = tcodperc + Val(t)
        ti = ti + Val(t)
        tt = tt + Val(t)
        msf1.AddItem F & Chr(9) & "" & Chr$(9) & rs2("cliente02") & Chr(9) & rs2("cuit02") & " " & Chr(9) & " " & nc & Chr(9) & "" & Chr(9) & t & Chr(9) & Format$(rs2("vta_02.num_int"), "00000")
        rs2.MoveNext

    Wend
    Set rs2 = Nothing
    
    If t <> "0" Then
        'totales por codigo
        msf1.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "======================"
        msf1.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "Total " & rs("descripcion") & Chr$(9) & "" & Chr(9) & Format$(ti, "########0.00")
        msf1.AddItem ""
    End If
    ti = 0
        
    rs.MoveNext
 Wend
  msf1.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "TOTAL PERCEPCIONES " & Chr$(9) & "" & Chr(9) & Format$(tt, "########0.00")
         
       




  
  
   
   
   
End Sub





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
msf1.TextMatrix(0, 2) = "Cliente"
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
Call carga_percepciones(c_concepto, "P")
c_concepto.AddItem "<Todos>", 0
c_concepto.ListIndex = 0


End Sub
Sub cargaret()
'impuestos
c_imp.clear
c_imp.AddItem "<Todos>", 0
c_imp.AddItem "Iva", 1
c_imp.AddItem "Ing.Brutos", 2
c_imp.AddItem "Ganancias", 3
c_imp.AddItem "Seg. Social", 4
c_imp.AddItem "Otros", 5
c_imp.ListIndex = 0

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
     Load vta_cc_detalle
     vta_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 7)
     vta_cc_detalle.Show
  
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
