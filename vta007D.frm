VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_recibo4 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INGRESO TRANSFERNCIAS"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3840
   ScaleWidth      =   7005
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   3495
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   6615
      Begin VB.TextBox t_modulo 
         Height          =   375
         Left            =   5520
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox c_banco 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   1800
         Width           =   4575
      End
      Begin VB.TextBox t_funcion 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   6000
         MaxLength       =   8
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox t_fechad 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox T_NUMCH 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox t_importe 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox t_titular 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   3
         Top             =   2280
         Width           =   4335
      End
      Begin VB.TextBox t_NUMINT 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Funcion"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4680
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Num.Int."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Fecha "
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Importe"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Banco"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Num.Transf."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   5
      Top             =   3810
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   53
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
            TextSave        =   "28/10/2010"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "07:12 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_recibo4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984



Private Sub c_banco_LostFocus()
If c_banco.ListIndex < 0 Then
  c_banco.ListIndex = 0
End If
End Sub

Private Sub Form_Activate()
Call limpia
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   
End Select
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 4)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call carga_formas_pago(c_banco, "B")
c_banco.ListIndex = 0
End Sub



 
  
Sub limpia()
T_NUMCH = ""
t_fechai = ""
t_fechad = ""
t_banco = ""
t_sucursal = ""
t_titular = ""
t_importe = ""
c_banco.ListIndex = 0
End Sub



Private Sub t_fechad_LostFocus()
If t_fechad <> "" Then
  If Not IsDate(t_fechad) Then
    t_fechad = Format$(Now, "dd/mm/yyyy")
  End If
Else
  t_fechad = Format$(Now, "dd/mm/yyyy")
End If
  
End Sub

Private Sub t_importe_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If verifica Then
     Call modificarenglon
     Call limpia
  Else
    MsgBox ("Error en los datos ingresados")
  End If
  Me.Hide
 Else
  Call solonum(KeyAscii, 1)
 End If
End Sub
Function verifica() As Boolean
 V = True
 If t_fechad = "" Then
    V = False
 Else
   If Not IsDate(t_fechad) Then
     V = False
   End If
 End If
 
 If Val(t_importe) <= 0 Then
   V = False
 End If
 
 verifica = V
 
End Function
Sub modificarenglon()
  Set rs = New ADODB.Recordset
  q = "select * from cyb_01  where [id_forma_pago] = " & c_banco.ItemData(c_banco.ListIndex)
  rs.MaxRecords = 1
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
    cta = rs("id_cuenta_cont")
  Else
    cta = 0
  End If
  Set rs = Nothing
  
  
  ip = Format$(4, "000")
  d = "Transf."
  i = Format$(Val(t_importe), "######0.00")
  Select Case t_modulo
  Case Is = "R"
   If t_fp <> "" Then
     r = Val(t_fp)
     vta_recibo.msf2.AddItem ip & Chr(9) & d & Chr(9) & Format$(Val(T_NUMCH), "0000000000") & Chr(9) & Left$(c_banco, 49) & " " & Chr(9) & Left$(t_tiular, 49) & " " & Chr(9) & Left$(vta_recibo.denominACION, 49) & " " & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fechad & Chr$(9) & c_banco.ItemData(c_banco.ListIndex) & Chr(9) & cta, r
     vta_recibo.msf2.RemoveItem r + 1
   Else
     vta_recibo.msf2.AddItem ip & Chr(9) & d & Chr(9) & Format$(Val(T_NUMCH), "0000000000") & Chr(9) & Left$(c_banco, 49) & " " & Chr(9) & Left$(t_titular, 49) & " " & Chr(9) & Left$(vta_recibo.denominACION, 49) & " " & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fechad & Chr$(9) & c_banco.ItemData(c_banco.ListIndex) & Chr(9) & cta
   End If
 Case Is = "F"
   If t_fp <> "" Then
     r = Val(t_fp)
     vta_formapago.msf2.AddItem ip & Chr(9) & d & Chr(9) & Format$(Val(T_NUMCH), "0000000000") & Chr(9) & Left$(c_banco, 49) & " " & Chr(9) & Left$(t_tiular, 49) & " " & Chr(9) & Left$(vta_recibo.denominACION, 49) & " " & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fechad & Chr$(9) & c_banco.ItemData(c_banco.ListIndex) & Chr$(9) & cta, r
     vta_formapago.msf2.RemoveItem r + 1
   Else
     vta_formapago.msf2.AddItem ip & Chr(9) & d & Chr(9) & Format$(Val(T_NUMCH), "0000000000") & Chr(9) & Left$(c_banco, 49) & " " & Chr(9) & Left$(t_titular, 49) & " " & Chr(9) & Left$(vta_recibo.denominACION, 49) & " " & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fechad & Chr$(9) & c_banco.ItemData(c_banco.ListIndex) & Chr$(9) & cta
   End If
 Case Is = "Q"
   If t_fp <> "" Then
     r = Val(t_fp)
     fsc_formapago.msf2.AddItem ip & Chr(9) & d & Chr(9) & Format$(Val(T_NUMCH), "0000000000") & Chr(9) & Left$(c_banco, 49) & " " & Chr(9) & Left$(t_tiular, 49) & " " & Chr(9) & Left$(vta_recibo.denominACION, 49) & " " & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fechad & Chr$(9) & c_banco.ItemData(c_banco.ListIndex) & Chr$(9) & cta & Chr(9) & "T.B. " & Format$(Val(T_NUMCH), "0000000000"), r
     fsc_formapago.msf2.RemoveItem r + 1
   Else
     fsc_formapago.msf2.AddItem ip & Chr(9) & d & Chr(9) & Format$(Val(T_NUMCH), "0000000000") & Chr(9) & Left$(c_banco, 49) & " " & Chr(9) & Left$(t_titular, 49) & " " & Chr(9) & Left$(vta_recibo.denominACION, 49) & " " & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fechad & Chr$(9) & c_banco.ItemData(c_banco.ListIndex) & Chr$(9) & cta & Chr(9) & "T.B. " & Format$(Val(T_NUMCH), "0000000000")
   End If
 End Select
 
End Sub

Private Sub T_NUMCH_LostFocus()
T_NUMCH = Format$(Val(T_NUMCH), "00000000")
End Sub


Private Sub t_titular_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF6 Then
 Select Case t_modulo
 Case Is = "R"
  t_titular = vta_recibo.denominACION
 Case Is = "F"
  t_titular = vta_facturacion.c_prov
 Case Is = "Q"
  t_titular = "Tique Contado"
 
 End Select
End If
End Sub

Private Sub t_titular_LostFocus()
t_titular = RTrim$(t_titular) & " "

End Sub
