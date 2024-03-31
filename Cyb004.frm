VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form cyb_cajadiaria 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CAJA DIARIA"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12045
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8520
   ScaleWidth      =   12045
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   25
      Top             =   6720
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "Cyb004.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "Cyb004.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Composicion a la Fecha (Composicion)"
      Height          =   2415
      Left            =   0
      TabIndex        =   22
      Top             =   5640
      Width           =   3975
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   3735
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   5760
      TabIndex        =   21
      Top             =   120
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   12632256
      Appearance      =   1
      StartOfWeek     =   114491393
      CurrentDate     =   38750
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Saldo Caja"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   4920
      TabIndex        =   12
      Top             =   6120
      Width           =   3735
      Begin VB.TextBox t_sc 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         TabIndex        =   20
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox t_s 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         TabIndex        =   19
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox t_e 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         TabIndex        =   18
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox t_si 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00800080&
         Caption         =   "Saldo Caja:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800080&
         Caption         =   "Salidas:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800080&
         Caption         =   "Entradas:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         Caption         =   "Saldo Anterior:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Saldo del Dia (Composicion)"
      Height          =   2535
      Left            =   7920
      TabIndex        =   10
      Top             =   3000
      Width           =   3975
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2160
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Salidas (Composicion)"
      Height          =   2535
      Left            =   3960
      TabIndex        =   8
      Top             =   3000
      Width           =   3975
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2160
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Entradas (Composicion)"
      Height          =   2535
      Left            =   0
      TabIndex        =   6
      Top             =   3000
      Width           =   3975
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2160
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha"
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3615
      Begin VB.TextBox t_fecha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   495
         Left            =   3120
         TabIndex        =   24
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Caja:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Saldo Anterior (Composicion)"
      Height          =   2055
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   3975
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3735
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   8160
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   635
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
            TextSave        =   "31/03/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:34 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cyb_cajadiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim sf As Double

Sub carga()
'genera saldo inicial
 List1.clear
 List2.clear
 List3.clear
 List4.clear
 List5.clear
 Call saldoi
 Call entradas
 Call saldof
 
End Sub

Sub saldoi()
 t = 0
 i = Space$(10)
 Set rs = New ADODB.Recordset
 q = "select * from cyb_01 where [caja] = 'S'"
 rs.Open q, cn1
 While Not rs.EOF
    Set rs1 = New ADODB.Recordset
    q = "select * from cyb_05 where [id_forma_pago] = " & rs("id_forma_pago") & " and datevalue([fecha]) < datevalue('" & t_fecha & "')"
    rs1.Open q, cn1
    si2 = 0
    While Not rs1.EOF
      If rs1("ubicacion") = "D" Then
         si2 = si2 + rs1("importe")
      Else
         si2 = si2 - rs1("importe")
      End If
      rs1.MoveNext
    Wend
    Set rs1 = Nothing
    RSet i = Format$(si2, "######0.00")
    List1.AddItem "[" & Format$(rs("id_forma_pago"), "000") & "] " & Format$(Left$(rs("descripcion"), 14), ">@@@@@@@@@@@@@@!") & "  $ " & i
    t = t + si2
    rs.MoveNext
 Wend
   RSet i = Format$(t, "######0.00")
   List1.AddItem "                      ----------------"
   List1.AddItem "                Total" & "  $ " & i
   t_si = Format$(t, "######0.00")
   
 Set rs = Nothing
End Sub
Sub entradas()
 te = 0
 ts = 0
 i = Space$(10)
 Set rs = New ADODB.Recordset
 q = "select * from cyb_01 where [caja] = 'S'"
 rs.Open q, cn1
 While Not rs.EOF
    Set rs1 = New ADODB.Recordset
    q = "select * from cyb_05 where [id_forma_pago] = " & rs("id_forma_pago") & " and datevalue([fecha]) = datevalue('" & t_fecha & "')"
    rs1.Open q, cn1
    e = 0
    s = 0
    sd = 0
    st = 0
    While Not rs1.EOF
       If rs1("UBICACION") = "D" Then
         e = e + rs1("importe")
         te = te + rs1("importe")
         sd = sd + rs1("importe")
       Else
         s = s - rs1("importe")
         ts = ts - rs1("importe")
         sd = sd - rs1("importe")
       End If
        
      rs1.MoveNext
    Wend
    Set rs1 = Nothing
    RSet i = Format$(e, "######0.00")
    List2.AddItem "[" & Format$(rs("id_forma_pago"), "000") & "] " & Format$(Left$(rs("descripcion"), 14), ">@@@@@@@@@@@@@@!") & "  $ " & i
    RSet i = Format$(s, "######0.00")
    List3.AddItem "[" & Format$(rs("id_forma_pago"), "000") & "] " & Format$(Left$(rs("descripcion"), 14), ">@@@@@@@@@@@@@@!") & "  $ " & i
    'agrega saldo final
    RSet i = Format$(sd, "######0.00")
    List4.AddItem "[" & Format$(rs("id_forma_pago"), "000") & "] " & Format$(Left$(rs("descripcion"), 14), ">@@@@@@@@@@@@@@!") & "  $ " & i
    st = st + sd
    rs.MoveNext
 Wend
   RSet i = Format$(te, "######0.00")
   List2.AddItem "                     ----------------"
   List2.AddItem "               Total" & "  $ " & i
   RSet i = Format$(ts, "######0.00")
   List3.AddItem "                     ----------------"
   List3.AddItem "               Total" & "  $ " & i
   
   RSet i = Format$(te - ts, "######0.00")
   
   List4.AddItem "                     ----------------"
   
   List4.AddItem "               Total" & "  $ " & i
   t_e = Format$(te, "######0.00")
   t_s = Format$(ts, "######0.00")

   t_sc = Format$(Val(t_si) + Val(t_e) + Val(t_s), "######0.00")
   
 Set rs = Nothing

End Sub


Sub saldof()
 te = 0
 i = Space$(10)
 Set rs = New ADODB.Recordset
 q = "select * from cyb_01 where [caja] = 'S'"
 rs.Open q, cn1
 While Not rs.EOF
    Set rs1 = New ADODB.Recordset
    q = "select * from cyb_05 where [id_forma_pago] = " & rs("id_forma_pago") & " and datevalue([fecha]) <= datevalue('" & t_fecha & "')"
    rs1.Open q, cn1
    sd = 0
    While Not rs1.EOF
       If rs1("UBICACION") = "D" Then
         sd = sd + rs1("importe")
         te = te + rs1("importe")
       Else
         sd = sd - rs1("importe")
         te = te - rs1("importe")
       End If
        
      rs1.MoveNext
    Wend
    Set rs1 = Nothing
    RSet i = Format$(sd, "######0.00")
    List5.AddItem "[" & Format$(rs("id_forma_pago"), "000") & "] " & Format$(Left$(rs("descripcion"), 14), ">@@@@@@@@@@@@@@!") & "  $ " & i
    rs.MoveNext
 Wend
   RSet i = Format$(te, "######0.00")
   List5.AddItem "                     ----------------"
   List5.AddItem "               Total" & "  $ " & i
   Set rs = Nothing

End Sub


Private Sub btnacepta_Click()
Call carga

End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Private Sub cal1_DblClick()
t_fecha = cal1
cal1.Visible = False
Call carga
End Sub

Private Sub cal1_LostFocus()
t_fecha = cal1
cal1.Visible = False
Call carga
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF12 Then
  gen_tools.Show
End If

If KeyCode = vbKeyF7 Then
  Call imprimir
End If

End Sub
Sub imprimir()
J = MsgBox("Prepare Impresora y Confirme", 4)

If J = 6 Then
    l1 = "---------------------------------"
    t1 = "                                        "
    Call imprimeempresa(14)
    Printer.FontSize = 12
    Printer.Print "COMPOSICION DE SALDOS DE CAJA AL DIA: " & t_fecha
    Printer.FontName = "Courier New"
    Printer.FontSize = 10
    Printer.Print
    Printer.Print "Saldo Anterior"
    Printer.Print l1
    Printer.Print "Concepto              Importe"
    Printer.Print l1
    g = 0
    While g < List1.ListCount
      Printer.Print List1.List(g)
      g = g + 1
    Wend
    Printer.Print
    Printer.Print t1 & "Entradas"
    Printer.Print t1 & l1
    Printer.Print t1 & "Concepto              Importe"
    Printer.Print t1 & l1
    g = 0
    While g < List2.ListCount
      Printer.Print t1 & List2.List(g)
      g = g + 1
    Wend
    Printer.Print
    Printer.Print t1 & "Salidas"
    Printer.Print t1 & l1
    Printer.Print t1 & "Concepto              Importe"
    Printer.Print t1 & l1
    g = 0
    While g < List3.ListCount
      Printer.Print t1 & List3.List(g)
      g = g + 1
    Wend
    Printer.Print
    Printer.Print "Composicion a la fecha"
    Printer.Print l1
    Printer.Print "Concepto              Importe"
    Printer.Print l1
    g = 0
    While g < List5.ListCount
      Printer.Print List5.List(g)
      g = g + 1
    Wend
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print "________________________________________________________________"
    Printer.Print "Fecha Imp." & Format$(Now, "dd/mm/yyyy") & "   Nro.Hoja: 1" & Format$(nh, "000") & "     Emitido por: " & glo.usuario

    
    Printer.EndDoc
 End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Unload Me
End If

End Sub

Private Sub Form_Load()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - [F12] Herramientas"
t_fecha = Format$(Now, "dd/mm/yyyy")
cal1.Visible = False
Call carga

End Sub

  




Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  FP = Val(Mid$(List1, 2, 3))
  If FP > 0 Then
    Load CYB_cc_detalle
    CYB_cc_detalle.t_idfp = FP
    CYB_cc_detalle.t_fp = Mid$(List1, 7, 20)
    CYB_cc_detalle.t_op = "A"
    CYB_cc_detalle.t_fecha = t_fecha
    CYB_cc_detalle.Show
  End If
End If
End Sub

Private Sub List1_LostFocus()
List1.ListIndex = -1
End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  FP = Val(Mid$(List2, 2, 3))
  If FP > 0 Then
    Load CYB_cc_detalle
    CYB_cc_detalle.t_idfp = FP
    CYB_cc_detalle.t_fp = Mid$(List2, 7, 20)
    CYB_cc_detalle.t_op = "E"
    CYB_cc_detalle.t_fecha = t_fecha
    CYB_cc_detalle.Show
  End If
End If

End Sub

Private Sub List2_LostFocus()
List2.ListIndex = -1
End Sub

Private Sub List3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  FP = Val(Mid$(List3, 2, 3))
  If FP > 0 Then
    Load CYB_cc_detalle
    CYB_cc_detalle.t_idfp = FP
    CYB_cc_detalle.t_fp = Mid$(List3, 7, 20)
    CYB_cc_detalle.t_op = "S"
    CYB_cc_detalle.t_fecha = t_fecha
    CYB_cc_detalle.Show
  End If
End If

End Sub

Private Sub List3_LostFocus()
List3.ListIndex = -1
End Sub

Private Sub List4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  FP = Val(Mid$(List4, 2, 3))
  If FP > 0 Then
    Load CYB_cc_detalle
    CYB_cc_detalle.t_idfp = FP
    CYB_cc_detalle.t_fp = Mid$(List4, 7, 20)
    CYB_cc_detalle.t_op = "D"
    CYB_cc_detalle.t_fecha = t_fecha
    CYB_cc_detalle.Show
  End If
End If

End Sub

Private Sub List4_LostFocus()
List4.ListIndex = -1
End Sub

Private Sub List5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  FP = Val(Mid$(List5, 2, 3))
  If FP > 0 Then
    Load CYB_cc_detalle
    CYB_cc_detalle.t_idfp = FP
    CYB_cc_detalle.t_fp = Mid$(List5, 7, 20)
    CYB_cc_detalle.t_op = "T"
    CYB_cc_detalle.t_fecha = t_fecha
    CYB_cc_detalle.Show
  End If
End If

End Sub

Private Sub t_fecha_DblClick()
cal1.Visible = True
End Sub

Private Sub t_fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = Format$(Now, "dd/mm/yyyy")
  End If
  Call carga
End If
  
End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
     t_fecha = Format$(t_fecha, "dd/mm/yyyy")
  End If
Else
  t_fecha = Format$(t_fecha, "dd/mm/yyyy")
End If
  
End Sub

Private Sub UpDown1_DownClick()
 t_fecha = DateValue(t_fecha) - 1
 Call carga
End Sub

Private Sub UpDown1_UpClick()
t_fecha = DateValue(t_fecha) + 1
Call carga

End Sub
