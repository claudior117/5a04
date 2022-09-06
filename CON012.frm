VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form con_busca_comp_apoc 
   BackColor       =   &H00E0E0E0&
   Caption         =   "BUSCADOR DE FACTURAS APÓCRIFAS"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   5040
      TabIndex        =   8
      Top             =   120
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   174915585
      CurrentDate     =   38754
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5895
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10398
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   4335
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1935
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
         Picture         =   "CON012.frx":0000
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
         Picture         =   "CON012.frx":0882
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
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   12347
            MinWidth        =   12347
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "05/09/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "01:28 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "con_busca_comp_apoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim saldoanterior As Double
Sub carga()
  Call armagrid
  espere.Show
  espere.Refresh
  q = "select * from fa order by [cuit]"
  Set rs = New ADODB.Recordset
  rs.Open q, cnib
  r = 0
  While Not rs.EOF
      r = r + 1
      espere.Label1 = "Validando..." & r
      espere.Label1.Refresh
      q = "select [fecha], [total], [cuit05], [letra], [sucursal], [num_comprobante], [num_int], [proveedor05] from a5 where [cuit05] = " & rs("cuit")
      c = " and "
      If t_fecha <> "" Then
       q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
       c = " and "
      End If
    
      If t_fecha2 <> "" Then
       q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
      End If
      q = q & " order by [fecha] "
    
      Set rs1 = New ADODB.Recordset
      rs1.Open q, cn1
      t = 0
      While Not rs1.EOF
         F = rs1("fecha")
         i = rs1("total")
         c = rs1("cuit05")
         cp = rs1("letra") & Format$(rs1("sucursal"), "0000") & "-" & Format$(rs1("num_comprobante"), "00000000")
         FP = rs("fecha_publicacion")
         ni = rs1("num_int")
         p = rs1("proveedor05")
         msf1.AddItem F & Chr(9) & cp & Chr(9) & c & Chr(9) & p & Chr(9) & i & Chr(9) & FP & Chr(9) & "" & Chr(9) & ni
         rs1.MoveNext
      Wend
      rs.MoveNext
     Wend
     Unload espere
  
End Sub



Private Sub btnacepta_Click()
J = MsgBox("Este proceso puede demorar, ¿esta segur que quiere continuar?", 4)
If J = 6 Then
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
  msf1.Cols = 8
  msf1.ColWidth(0) = 1200
  msf1.ColWidth(1) = 2500
  msf1.ColWidth(2) = 1200
  msf1.ColWidth(3) = 3500
  msf1.ColWidth(4) = 1200
  msf1.ColWidth(5) = 1200
  msf1.ColWidth(6) = 1200
  msf1.ColWidth(7) = 1200
  msf1.TextMatrix(0, 0) = "Fecha"
  msf1.TextMatrix(0, 1) = "Comprobante"
  msf1.TextMatrix(0, 2) = "Cuit"
  msf1.TextMatrix(0, 3) = "Proveedor"
  msf1.TextMatrix(0, 4) = "Importe"
  msf1.TextMatrix(0, 5) = "Fecha Publicado"
  msf1.TextMatrix(0, 6) = "Obs."
  msf1.TextMatrix(0, 7) = "NI"
  
  For i = 0 To 1
    msf1.ColAlignment(i) = 1
  Next i
  For i = 2 To 6
    msf1.ColAlignment(i) = 9
  Next i
 
End Sub








Private Sub cal1_DblClick()
t_fecha = cal1
cal1.Visible = False
End Sub

Private Sub cal1_LostFocus()
t_fecha = cal1
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
    Call TabEnter2(Me, 2)
  'Case Is = 27
  '      Me.Hide
End Select

End Sub

Private Sub Form_Load()


Call armagrid
Call barraesag(Me)
cal1.Visible = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Unload vta_clientes
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - [F11] Excel  "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Call nivel_acceso(1)
  If para.id_grupo_modulo_actual >= 4 Then
    J = MsgBox("Prepare Impresora y Confirme", 4)
    If J = 6 Then
     Dim c(15) As Double

      c(0) = 0
      c(1) = 1
      c(2) = 2
      c(3) = 3
      c(4) = 4
      c(5) = 5
      c(6) = 6
      c(7) = 7
      
      For i = 8 To 14
        c(i) = -1
      Next i
     End If
     Call imprimegrid(msf1, c(), "COMPROBANTES APOCRIFOS INGRESADOS", "", " ", "Periodo: " & t_fecha & "  " & t_fecha2, 85, 7, True, False)

         
  End If
  
End If


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
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
