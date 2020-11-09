VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form con_ley23966A 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UTILIZACION SALDO SUBSIDIO en DECLARACION JURADA"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2370
   ScaleWidth      =   7005
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6615
      Begin VB.TextBox t_fechad 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_importe 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox t_titular 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   1
         Top             =   720
         Width           =   4935
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Fecha Cierre periodo ddjj"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Importe"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   3
      Top             =   2340
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
            TextSave        =   "30/03/2011"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "05:34 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "con_ley23966A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984




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
    Call TabEnter2(Me, 2)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
End Sub



 
  
Sub limpia()
t_fechad = ""
t_titular = ""
t_importe = ""
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
     J = MsgBox("Acepta grabar Operacion", 4)
     If J = 6 Then
       Call graba
       Call con_ley23966.carga
       Unload Me
     End If
  Else
    MsgBox ("Error en los datos ingresados")
  End If
  Me.Hide
 Else
  Call solonum(KeyAscii, 1)
 End If
End Sub
Sub graba()
numint = saca_ultnumero_int_comp("C")

QUERY = "insert into a19([num_int], [litros], [pu_impuesto_int], [importe], [ubicacion], [detalle], [fecha])"
QUERY = QUERY & " VALUES (" & numint & ", 0, 0, " & Val(t_importe) & ", 'H', '" & Left$(t_titular, 50) & "', '" & t_fechad & "')"
cn1.Execute QUERY

End Sub
Function verifica() As Boolean
 V = True
 If t_fechad = "" Then
    MsgBox ("Fecha Incorrecta")
    V = False
 Else
   If Not IsDate(t_fechad) Then
     MsgBox ("Fecha Incorrecta")
     V = False
   End If
 End If
 
'verifica importe con saldo
If V = True Then
 If Val(t_importe) <= 0 Then
   MsgBox ("El importe no puede ser menor a 0")
   V = False
 Else
   q = "select * from a19 where datevalue([fecha]) <= datevalue('" & t_fechad & "')"
   Set rs = New adodb.Recordset
   rs.Open q, cn1
   s = 0
   While Not rs.EOF
     If rs("ubicacion") = "D" Then
       s = s + rs("importe")
     Else
       s = s - rs("importe")
     End If
     rs.MoveNext
   Wend
   Set rs = Nothing
   If Val(Format$(s, "#####0.00")) < Val(t_importe) Then
     V = False
     MsgBox ("El importe a aplicar no puede superar el saldo acumulado del subsidio")
   End If
   
 End If
End If
 verifica = V
 
End Function




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
