VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form com_vencimientos 
   BackColor       =   &H00E0E0E0&
   Caption         =   "VENCIMIENTOS"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   360
      TabIndex        =   16
      Top             =   840
      Width           =   6975
      Begin VB.ComboBox c_zona 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         Caption         =   "Zona:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Orden"
      Height          =   735
      Left            =   3960
      TabIndex        =   12
      Top             =   120
      Width           =   3375
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "fecha vto."
         Height          =   195
         Left            =   1560
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   4080
      TabIndex        =   10
      Top             =   2040
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   179699713
      CurrentDate     =   38803
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fecha Vto. Desde - Hasta"
      Height          =   735
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Width           =   3495
      Begin VB.TextBox t_fecha2 
         Height          =   330
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox t_fecha 
         Height          =   330
         Left            =   120
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Moneda"
      Height          =   735
      Left            =   7440
      TabIndex        =   6
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton O_dolares 
         BackColor       =   &H00E0E0E0&
         Caption         =   "U$s"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Tag             =   "D"
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton O_pesos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pesos"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Tag             =   "P"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10080
      TabIndex        =   3
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "com005.frx":0000
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
         Picture         =   "com005.frx":0882
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
            TextSave        =   "02/04/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "04:27 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   4695
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8281
      _Version        =   393216
      AllowUserResizing=   1
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
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   10440
      TabIndex        =   15
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "com_vencimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String
Private saldoant As Double
Private saldoact As Double
'FIXIT: Declare 'saf' and 'df' and 'hf' and 'sf' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim saf, df, hf, sf, sof As Double




Private Sub btnacepta_Click()
    Call carga
End Sub

Sub carga()
espere.Show
  espere.Label1 = "Cargando comprobantes emitidos...."
  espere.Refresh
  Call armagrid
  q = "select * from a5, g2 where [id_tipocomp] = [id_tipo_comp]   and a5.[compra] <> 'N' and [saldo_impago] > 0"
  c = " and "
  
  
  If t_fecha <> "" And IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha_vto]) >= datevalue('" & t_fecha & "')"
  End If
  
  If t_fecha2 <> "" And IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha_vto]) <= datevalue('" & t_fecha2 & "')"
  End If
    
   
   If c_zona.ListIndex > 0 Then
    q = q & c & " [zona] = " & c_zona.ItemData(c_zona.ListIndex)
   End If

 
 If Option1 = True Then
    q = q & " order by [proveedor05], [fecha_vto]"
 Else
    q = q & " order by [fecha_vto], [proveedor05]"
 End If
 
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  t = 0
  tsi = 0
  reg = 0
  While Not rs.EOF
     F = rs("fecha_vto")
     nc = rs("abreviatura") & " " & rs("letra") & " " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
     d = Format$(rs("total"), "######0.00")
     p = rs("proveedor05")
     ni = rs("num_int")
     If O_pesos Then
       m = "P"
     Else
       m = "D"
     End If
     If m = rs("moneda") Then
       d = Format$(rs("total"), "######0.00")
       t = t + Val(d)
      Else
       If m = "P" Then
         'informe en p y comp en dolares
         d = Format$(rs("total") * rs("cotiz_dolar"), "######0.00")
         t = t + Val(d)
       Else
          'informe en d y comp en p
         d = Format$(rs("total") / rs("cotiz_dolar"), "######0.00")
         t = t + Val(d)
       End If
     End If
     
     
    If m = "P" Then
       si = Format$(rs("saldo_impago"), "######0.00")
       tsi = tsi + Val(si)
    Else
       si = Format$(rs("saldo_impago") / rs("cotiz_dolar"), "######0.00")
       tsi = tsi + Val(si)
    End If
     
     msf1.AddItem F & Chr(9) & p & Chr(9) & nc & Chr(9) & d & Chr(9) & si & Chr(9) & rs("num_int")
     reg = reg + 1
     Label5 = reg
     Label5.Refresh
    rs.MoveNext
  Wend
  msf1.AddItem ""
  msf1.AddItem "" & Chr(9) & "Comprobantes: " & reg & Chr(9) & "" & Chr(9) & Format$(t, "#####0.00") & Chr(9) & Format$(tsi, "#####0.00")
  Unload espere

End Sub





Private Sub btnsale_Click()

Unload Me
End Sub





Private Sub c_zona_LostFocus()
If c_zona.ListIndex < 0 Then
  c_zona.ListIndex = 0
End If
End Sub



Private Sub cal1_LostFocus()
cal1.Visible = False
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
End Select

End Sub
Sub armagrid()
'armar grilla
  msf1.clear
  msf1.Rows = 1
  msf1.Cols = 6
  If O_pesos = True Then
    tm = "$ "
  Else
    tm = "U$s "
  End If
  
  msf1.ColWidth(0) = 1200
  msf1.ColWidth(1) = 4000
  msf1.ColWidth(2) = 2200
  msf1.ColWidth(3) = 1400
  msf1.ColWidth(4) = 1400
  msf1.ColWidth(5) = 800
  msf1.TextMatrix(0, 0) = "Fecha Vto."
  msf1.TextMatrix(0, 1) = "Proveedor"
  msf1.TextMatrix(0, 2) = "Comprobante"
  msf1.TextMatrix(0, 3) = tm & "Total"
  msf1.TextMatrix(0, 4) = tm & "Pend."
  msf1.TextMatrix(0, 5) = "Nro.Int."
  
  
  For i = 0 To 2
    msf1.ColAlignment(i) = 1 'izq
  Next i
  For i = 3 To 5
   msf1.ColAlignment(1) = 9 'der
  Next i
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 2)
  
End Select


End Sub

Private Sub Form_Load()
Call barra(Me)

If para.moneda = "P" Then
  O_pesos = Checked
Else
  O_dolares = Checked
End If


Call carga_zonas(c_zona)
c_zona.AddItem "<Todas>", 0
c_zona.ListIndex = 0

cal1.Visible = False
Call armagrid

Option1 = True
Option4 = True
End Sub




Private Sub cal1_DblClick()
If cal1.Tag = "1" Then
  t_fecha = cal1
  t_fecha.SetFocus
Else
  t_fecha2 = cal1
  t_fecha2.SetFocus
End If
cal1.Visible = False
End Sub



Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - [F11] Excel - [ENTER] Visualiza "
End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    If O_pesos Then
      m = "Pesos ($)"
    Else
      m = "Dolares (U$s)"
    End If
    
    If c_vend.ListIndex > 0 Then
       v = "   Vendedor: " & c_vend
    Else
       v = " "
    End If
      
      c(0) = 0
      c(1) = 1
      c(2) = 2
      c(3) = 3
      c(4) = 4
      c(5) = 5
      
      For i = 6 To 14
        c(i) = -1
      Next i
      Call imprimegrid(msf1, c(), Space$(40) & "VENCIMIENTOS PENDIENTES", "   Periodo...: " & t_fecha & " - " & t_fecha2, "   Moneda....: " & m, v, 70, 8, True, True)
  End If

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
cal1.Tag = 1
cal1.SetFocus
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

Private Sub t_fecha2_DblClick()
cal1.Visible = True
cal1.Tag = 2
cal1.SetFocus

End Sub

Private Sub t_fecha2_GotFocus()
t_fecha2 = ""
End Sub

'FIXIT: t_fecha2_LinkOpen event no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
Private Sub t_fecha2_LinkOpen(Cancel As Integer)

End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
  If Not IsDate(t_fecha2) Then
    t_fecha2 = ""
  End If
End If
End Sub
