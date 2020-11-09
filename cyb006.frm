VERSION 5.00
Begin VB.Form cyb_generachpropios 
   Caption         =   "Genera Cheques Propios"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   4200
      TabIndex        =   6
      Top             =   2280
      Width           =   2655
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generar"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Generacion de Chequeras"
      Height          =   2055
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   6495
      Begin VB.TextBox t_cantch 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox t_chequera 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox t_ch2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox t_ch1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   1
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox c_banco 
         Height          =   315
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Cant. Cheques:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5040
         TabIndex        =   13
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Nro. Chequera"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Nro. Ch. Final"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Nro. Ch. Inicial"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Cuenta Bancaria"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "cyb_generachpropios"
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

Private Sub Command1_Click()
 J = MsgBox("Confirma Generacion de Chequera", 4)
 If J = 6 Then
  Call graba
  Unload Me
 End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 3)
  Case Is = 27
        Me.Hide
End Select
End Sub
Sub graba()
'On Error GoTo errg
cn1.BeginTrans
For i = Val(t_ch1) To Val(t_ch2)
  Set rs = New ADODB.Recordset
  q = "select * from cyb_02 where [id_banco] = " & c_banco.ItemData(c_banco.ListIndex) & " and [num_cheque] = " & i
  rs.Open q, cn1
  If Not rs.BOF And Not rs.EOF Then
      Set rs = Nothing
  Else
     Set rs = Nothing
     QUERY = "INSERT INTO cyb_02([id_banco], [num_cheque], [fecha_emision], [fecha_dif], [estado], [destino], [importe], [num_mov_banco], [id_chequera], [num_int_op])"
     QUERY = QUERY & " VALUES (" & c_banco.ItemData(c_banco.ListIndex) & ", " & i & ", '" & Format$(Now, "dd/mm/yyyy") & "', '" & Format$(Now, "dd/mm/yyyy") & "', 'P', 'Pendiente', 0, 0, " & Val(t_chequera) & ", 0)"
     cn1.Execute QUERY
  End If
Next i
cn1.CommitTrans

Exit Sub

errg:
  MsgBox ("Error al generar chequera. Verifique los datos y reintente")
  cn1.RollbackTrans
  Exit Sub
End Sub
Private Sub Form_Load()
Call INICIALIZA2(Me)
Call carga_formas_pago(c_banco, "B")
c_banco.ListIndex = 0

End Sub

Private Sub t_ch1_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)

End Sub


Private Sub t_ch1_LostFocus()
If Val(t_ch1) <= 0 Then
  t_ch1 = 1
End If
End Sub

Private Sub t_ch2_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)

End Sub

Private Sub t_ch2_LostFocus()
If Val(t_ch1) <= 0 Then
  t_ch1 = 1
End If
t_cantch = Val(t_ch2) - Val(t_ch1) + 1
End Sub

Private Sub t_chequera_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)

End Sub
