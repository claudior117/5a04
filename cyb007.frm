VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form cyb_chpropios 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADMINISTRADOR CHEQUES PROPIOS"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8700
   ScaleWidth      =   11970
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   6960
      TabIndex        =   27
      Top             =   1680
      Width           =   2535
      Begin VB.TextBox t_chequera 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Chequera Nro:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Estado"
      Height          =   1335
      Left            =   240
      TabIndex        =   21
      Top             =   1080
      Width           =   6135
      Begin VB.ComboBox c_concilia 
         Height          =   315
         ItemData        =   "cyb007.frx":0000
         Left            =   1440
         List            =   "cyb007.frx":000D
         TabIndex        =   32
         Top             =   600
         Width           =   3975
      End
      Begin VB.ComboBox c_banco 
         Height          =   315
         ItemData        =   "cyb007.frx":003E
         Left            =   1440
         List            =   "cyb007.frx":005A
         TabIndex        =   24
         Text            =   "c_banco"
         Top             =   960
         Width           =   3975
      End
      Begin VB.ComboBox c_estados 
         Height          =   315
         ItemData        =   "cyb007.frx":00C1
         Left            =   1440
         List            =   "cyb007.frx":00DD
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label11 
         BackColor       =   &H00800000&
         Caption         =   "Conciliacion"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800000&
         Caption         =   "Banco"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800000&
         Caption         =   "Estado"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ordenado por"
      Height          =   615
      Left            =   6960
      TabIndex        =   18
      Top             =   960
      Width           =   4215
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Difereida"
         Height          =   255
         Left            =   2040
         TabIndex        =   20
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Numero "
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   15
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "cyb007.frx":014E
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "cyb007.frx":09D0
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   6960
      TabIndex        =   12
      Top             =   120
      Width           =   4215
      Begin VB.TextBox t_cedido 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   14
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Cedido a :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Diferida"
      Height          =   975
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   3015
      Begin VB.TextBox t_fecha3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_fecha4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H00800000&
         Caption         =   "Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00800000&
         Caption         =   "Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Ingreso"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      Begin VB.TextBox t_fecha2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox t_fecha1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H00800000&
         Caption         =   "Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00800000&
         Caption         =   "Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8445
      Width           =   11970
      _ExtentX        =   21114
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
            TextSave        =   "29/04/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "04:27 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   3480
      TabIndex        =   11
      Top             =   0
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   115539969
      CurrentDate     =   38754
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   4695
      Left            =   240
      TabIndex        =   26
      Top             =   2520
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8281
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   10080
      TabIndex        =   31
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000FF&
      Caption         =   "ESTADOS: ""P"" Pendiente(En Chequera) - ""E"" Entregados - ""A"" Anulado - ""T"" Depositado - ""J"" Cobrado - ""D"" Devuelto - ""V"" Vendidos."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   30
      Top             =   7560
      Width           =   9855
   End
End
Attribute VB_Name = "cyb_chpropios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Private Sub btnacepta_Click()
Call carga
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Private Sub c_banco_LostFocus()
If c_banco.ListIndex < 0 Then
  If Val(c_banco) > 0 Then
    c_banco.ListIndex = buscaindice(c_banco, Val(c_banco))
  Else
    c_banco.ListIndex = 0
  End If
End If
End Sub

Private Sub cal1_DblClick()
  Select Case cal1.Tag
    Case Is = "1"
      t_fecha1 = cal1
    Case Is = "2"
      t_fecha2 = cal1
    Case Is = "3"
      t_fecha3 = cal1
    Case Is = "4"
      t_fecha4 = cal1
   End Select
   cal1.Visible = False
  
End Sub

Private Sub cal1_LostFocus()
 cal1.Visible = False
End Sub

Private Sub Form_Activate()
cal1.Visible = False
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 10
msf1.ColWidth(0) = 1200
msf1.ColWidth(1) = 1200
msf1.ColWidth(2) = 2700
msf1.ColWidth(3) = 800
msf1.ColWidth(4) = 2700
msf1.ColWidth(5) = 1500
msf1.ColWidth(6) = 800
msf1.ColWidth(7) = 800
msf1.ColWidth(8) = 2000
msf1.ColWidth(9) = 1000
msf1.TextMatrix(0, 0) = "Nro."
msf1.TextMatrix(0, 1) = "Fecha Dif."
msf1.TextMatrix(0, 2) = "Banco"
msf1.TextMatrix(0, 3) = "Estado"
msf1.TextMatrix(0, 4) = "Destino"
msf1.TextMatrix(0, 5) = "Importe"
msf1.TextMatrix(0, 6) = "Entro?"
msf1.TextMatrix(0, 7) = "Chequera"
msf1.TextMatrix(0, 8) = "Nro. OP"
msf1.TextMatrix(0, 9) = "Id.Cuenta"
For i = 0 To 8
    msf1.ColAlignment(i) = 1
Next i
msf1.ColAlignment(5) = 7


End Sub


Private Sub carga()
Call armagrid

Set rs = New ADODB.Recordset

q = "select * from cyb_02, cyb_01 where [id_banco] = [id_forma_pago] "
c = " and "

If t_fecha1 <> "" And IsDate(t_fecha1) Then
  q = q & c & " datevalue([fecha_emision]) >= datevalue('" & t_fecha1 & "')"
  c = " and "
End If

If t_fecha2 <> "" And IsDate(t_fecha2) Then
  q = q & c & " datevalue([fecha_emision]) <= datevalue('" & t_fecha2 & "')"
  c = " and "
End If

If t_fecha3 <> "" And IsDate(t_fecha3) Then
  q = q & c & " datevalue([fecha_dif]) >= datevalue('" & t_fecha3 & "')"
  c = " and "
End If

If t_fecha4 <> "" And IsDate(t_fecha4) Then
  q = q & c & " datevalue([fecha_dif]) <= datevalue('" & t_fecha4 & "')"
  c = " and "
End If


If t_cedido <> "" Then
  q = q & c & " [destino] like '%" & t_cedido & "%'"
  c = " and "
End If

If c_estados.ListIndex > 0 Then
   q = q & c & " [estado] = '" & Mid$(c_estados, 1, 1) & "'"
   c = " and "

End If

If c_banco.ListIndex > 0 Then
   q = q & c & " [id_banco] = " & c_banco.ItemData(c_banco.ListIndex)
   c = " and "

End If

If Val(t_chequera) > 0 Then
   q = q & c & " [id_chequera] = " & Val(t_chequera)
End If




If Option1 = True Then
   q = q & " order by [num_cheque]"
Else
   q = q & " order by [fecha_dif]"
End If

rs.Open q, cn1
Set rs1 = New ADODB.Recordset
ich = 0
cch = 0
While Not rs.EOF
     Label6 = cch
     Label6.Refresh
     If rs("num_int_op") > 0 Then
        Set rs1 = New ADODB.Recordset
        q = "select * from a5 where [num_int] = " & rs("num_int_op")
        rs1.Open q, cn1
        If Not rs1.EOF And Not rs1.BOF Then
           nc = Format$(rs1("sucursal"), "0000") & "-" & Format$(rs1("num_comprobante"), "00000000")
        Else
           nc = "No Existe"
        End If
        Set rs1 = Nothing
     Else
       nc = ""
     End If
     
     m = "S"
     e = " "
     If rs("num_mov_banco") > 0 Then
        Set rs1 = New ADODB.Recordset
        q = "select [entro] from cyb_04 where [num_mov_banco] = " & rs("num_mov_banco")
        rs1.Open q, cn1
        If Not rs1.EOF And Not rs1.BOF Then
           e = rs1("Entro")
        Else
           e = "X"
        End If
        Set rs1 = Nothing
        If c_concilia.ListIndex > 0 Then
         If c_concilia.ListIndex = 1 Then
          If e = "S" Then
           m = "S"
          Else
           m = "N"
          End If
         Else
          If e = "N" Then
           m = "S"
          Else
           m = "N"
          End If
         End If
       End If
    End If
     
     If m = "S" Then
       ich = ich + rs("importe")
       cch = cch + 1
       msf1.AddItem Format$(rs("num_cheque"), "0000000000") & Chr$(9) & Format$(rs("fecha_dif"), "dd/mm/yyyy") & Chr$(9) & rs("descripcion") & Chr$(9) & rs("estado") & Chr$(9) & rs("destino") & Chr$(9) & Format$(rs("importe"), "#####0.00") & Chr$(9) & e & Chr$(9) & rs("id_chequera") & Chr$(9) & nc & Chr$(9) & Format$(rs("id_banco"), "000")
     End If
     rs.MoveNext
Wend
 msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "________________________"
 msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "Cant. Cheques: " & cch & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & Format$(ich, "#####0.00")


Set rs = Nothing
End Sub

Private Sub Form_Load()
Call INICIALIZA2(Me)
Call barraesag(Me)
Call armagrid
Option1 = True
c_estados.ListIndex = 1
Call carga_formas_pago(c_banco, "B")
c_banco.AddItem "<Todos>", 0
c_banco.ListIndex = 0
c_concilia.ListIndex = 0
End Sub

  

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[ENTER] Administrador Cheque - [F7] Imprime - [F11] Excel"

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
    Call imprimegrid(msf1, c(), "LISTADO CHEQUES PROPIOS", "Banco: " & c_banco, "Fecha Emision...: " & t_fecha1 & "**" & t_fecha2 & "       " & "Fecha Diferida...: " & t_fecha3 & "**" & t_fecha4, "Estado..:" & c_estados, 50, 9, True, False, "H")
  End If
    
End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If
End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 
 Call nivel_acceso(4)
 If para.id_grupo_modulo_actual >= 7 Then
  F = msf1.Row
  If F >= 1 Then
     Load cyb_chpropios2
     Set rs = New ADODB.Recordset
     q = "select * from cyb_02, cyb_01 where [id_banco] = " & Val(msf1.TextMatrix(F, 9)) & " and [num_cheque] = " & Val(msf1.TextMatrix(F, 0)) & " and [id_banco] = [id_forma_pago] "
     'MsgBox (q)
     rs.Open q, cn1
     If Not rs.EOF And Not rs.BOF Then
      
       cyb_chpropios2.t_ch = Format$(rs("num_cheque"), "0000000000")
       cyb_chpropios2.t_idbanco = rs("id_banco")
       cyb_chpropios2.t_banco = rs("descripcion")
       cyb_chpropios2.t_chequera = rs("id_chequera")
       cyb_chpropios2.t_estado = rs("estado")
       'debo buscar movimiento
        Set rs1 = New ADODB.Recordset
        q = "select * from cyb_04 where [num_mov_banco] = " & rs("num_mov_banco")
        rs1.Open q, cn1
        If Not rs1.BOF And Not rs1.EOF Then
          cyb_chpropios2.t_fecha = rs1("fecha")
          cyb_chpropios2.t_fechadif = rs1("fecha_dif")
          cyb_chpropios2.t_importe = rs1("importe")
          cyb_chpropios2.t_destino = rs("destino")
          cyb_chpropios2.t_numint = rs1("NUM_MOV_BANCO")
        End If
        Set rs1 = Nothing
        cyb_chpropios2.Show
     End If
     Set rs = Nothing
     
  End If
 Else
  Call sinpermisos
 End If
End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
End Sub

Private Sub t_cedido_GotFocus()
t_cedido = ""
End Sub

Private Sub t_fecha1_DblClick()
cal1.Visible = True
cal1.Tag = "1"
End Sub

Private Sub t_fecha2_DblClick()
cal1.Visible = True
cal1.Tag = "2"

End Sub

Private Sub t_fecha3_DblClick()
cal1.Visible = True
cal1.Tag = "3"

End Sub

Private Sub t_fecha3_GotFocus()
t_fecha3 = ""
End Sub

Private Sub t_fecha4_DblClick()
cal1.Visible = True
cal1.Tag = "4"

End Sub

Private Sub t_fecha4_GotFocus()
t_fecha4 = ""
End Sub
