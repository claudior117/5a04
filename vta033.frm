VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form vta_verviajes 
   BackColor       =   &H00E0E0E0&
   Caption         =   "CONTROL DE VIAJES POR TRASPORTE y CAMION (Solo Remitos de Venta)"
   ClientHeight    =   8745
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   12165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordenados por:"
      Height          =   615
      Left            =   240
      TabIndex        =   18
      Top             =   7320
      Width           =   3015
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Camion"
         Height          =   255
         Left            =   1560
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5055
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8916
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   2055
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   11535
      Begin VB.TextBox t_dominio 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   21
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox c_transp 
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Text            =   "c_transp"
         Top             =   240
         Width           =   4815
      End
      Begin VB.TextBox T_chofer 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   8280
         MaxLength       =   10
         TabIndex        =   12
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Text            =   "c_prov"
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Dominio:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   9960
         TabIndex        =   16
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Transporte:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Chofer:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6720
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Camion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   600
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
         Picture         =   "vta033.frx":0000
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
         Picture         =   "vta033.frx":0882
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
      Top             =   8490
      Width           =   12165
      _ExtentX        =   21458
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
            TextSave        =   "09/08/2010"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:15 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800080&
      Caption         =   "Transporte:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "vta_verviajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer


Sub carga()
    
  Call armagrid
  q = "select * from vta_02, a1, a17 where [vta_02.id_transporte] = [id_proveedor] and  [id_camion02] = [id_camion] and [id_tipocomp] = 45"
  c = " and "
  If c_prov.ListIndex > 0 Then
     q = q & c & " [id_camion02] = " & c_prov.ItemData(c_prov.ListIndex)
  End If
  
    If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  
  If c_transp.ListIndex > 0 Then
     q = q & c & " [vta_02.Id_transporte] = " & c_transp.ItemData(c_transp.ListIndex)
  End If
  
  If T_chofer <> "" Then
   q = q & c & " [chofer02] like '%" & T_chofer & "%'"
  End If
  
  If t_dominio <> "" Then
   q = q & c & " [dominio02] like '%" & t_dominio & "%'"
  End If
  
  
  If Option1 = True Then
    q = q & " order by [fecha], [num_comp]"
  Else
    q = q & " order by [camion02], [fecha], [num_comp]"
  End If
  
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  reg = 0
  While Not rs.EOF
     f = rs("fecha")
     nc = rs("letra") & " " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
     d = Format$(rs("total"), "######0.00")
     p = rs("cliente02")
     c = rs("camion")
     t = rs("denominacion")
     ch = rs("chofer02")
     ni = rs("num_int")
     
  msf1.AddItem f & Chr(9) & nc & Chr(9) & t & Chr(9) & c & Chr$(9) & ch & Chr(9) & p & Chr(9) & "" & Chr(9) & rs("num_int")
  reg = reg + 1
  Label9 = reg
  Label9.Refresh
  rs.MoveNext
 Wend

  msf1.AddItem ""
  msf1.AddItem "" & Chr(9) & "" & Chr(9) & "Comprobantes: " & reg '& chr (9) & "" & Chr(9) & "" & Chr(9) & "Totales:" & Chr(9) & Format$(t, "#####0.00") & Chr(9) & Format$(tfa, "#####0.00") & Chr(9) & Format$(tpe, "#####0.00") & Chr(9) & ""

  
 
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
msf1.Cols = 9
msf1.ColWidth(0) = 1200
msf1.ColWidth(1) = 1000
msf1.ColWidth(2) = 2500
msf1.ColWidth(3) = 2500
msf1.ColWidth(4) = 2500
msf1.ColWidth(5) = 2000
msf1.ColWidth(6) = 800
msf1.ColWidth(7) = 800

msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = "Remito"
msf1.TextMatrix(0, 2) = "Transporte"
msf1.TextMatrix(0, 3) = "Camion"
msf1.TextMatrix(0, 4) = "Chofer"
msf1.TextMatrix(0, 5) = "Cliente"
msf1.TextMatrix(0, 6) = ""
msf1.TextMatrix(0, 7) = "Nro. Int."
For i = 0 To 5
    msf1.ColAlignment(i) = 1 'izq
Next i
msf1.ColAlignment(6) = 9 'der

End Sub










Private Sub c_transp_LostFocus()
If c_transp.ListIndex < 0 Then
  c_transp.ListIndex = 0
End If
Call cargacamion
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub
Sub cargacamion()
 If c_transp.ListIndex = 0 Then
    ct = 0
 Else
   ct = c_transp.ItemData(c_transp.ListIndex)
 End If
 Call carga_camiones(c_prov, ct)
 c_prov.AddItem "<Todos>", 0
 c_prov.ListIndex = 0
End Sub

Private Sub Form_Load()

Call carga_transporte(c_transp)
c_transp.AddItem "<Todos>", 0
c_transp.ListIndex = 0

Option1 = True

Call cargacamion
Call armagrid
Call barraesag(Me)


End Sub


Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime - [ENTER] Detalla - [F11] Excel"
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
    c(0) = 6
    c(1) = 0
    c(2) = 1
    c(3) = 2
    c(4) = 3
    c(5) = 4
    c(6) = 5
    For i = 7 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "CONTROL DE VIAJES", "Transp:" & c_transp & "           Camion: " & c_prov, "Fecha desde: " & t_fecha & "  Fecha hasta: " & t_fecha2, "Chofer: " & c_chofer, 45, 8, True, False, "H")
  End If

End If


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

End Sub


Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load vta_cc_detalle
    vta_cc_detalle.t_prov = msf1.TextMatrix(msf1.Row, 5)
    vta_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 7)
    vta_cc_detalle.Show
  End If
End If

End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight

End Sub

Private Sub T_chofer_GotFocus()
T_chofer = ""
End Sub


Private Sub t_dominio_GotFocus()
t_dominio = ""

End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha2_GotFocus()
t_fecha2 = ""
End Sub

