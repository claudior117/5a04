VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form cyb_concilia 
   Caption         =   "Conciliacion Bancaria"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox t_idbanco 
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox t_tipomov 
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Caption         =   "Funciones"
      Height          =   855
      Left            =   2640
      TabIndex        =   10
      Top             =   2520
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   840
         Picture         =   "cyb_018.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton btnacepta 
         Height          =   495
         Left            =   120
         Picture         =   "cyb_018.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Estado Conciliacion"
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   4095
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox t_fecha 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox t_entro 
         Height          =   285
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   0
         Top             =   360
         Width           =   495
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         BackColor       =   &H000000FF&
         Caption         =   "Entro:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000FF&
         Caption         =   "Entro:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      Begin VB.TextBox t_mov 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox t_id 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Importe:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Movimiento Nro:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "cyb_concilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Private Sub btnacepta_Click()
If verificaperiodog(t_fecha) = "A" Then
q = "select * from cyb_04 where [num_mov_banco] = " & Val(t_id)
Set rs = New ADODB.Recordset
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
If Not rs.EOF And Not rs.BOF Then
   rs("entro") = t_entro
   rs("fecha_acreed") = t_fecha
   rs.Update
End If
Set rs = Nothing

If t_tipomov = 1 Or t_tipomov = 50 Or t_tipomov = 80 Then
   'limpio asiento de conciliacion si existia
   'siempre hay dos asientos con cheque diferidos , uno cuando se emite y otro cuando ingresa
   q = "select * from c_02 where [modulo] = 'B' and [num_mov_int] = " & Val(t_id) & " and [descripcion] like '%Conciliacion%'"
   Set rs = New ADODB.Recordset
   rs.MaxRecords = 1
   rs.Open q, cn1
   If Not rs.EOF And Not rs.BOF Then
        nicgr = rs("num_interno")
   Else
        nicgr = 0
   End If
   Set rs = Nothing
      
   cn1.BeginTrans
   QUERY = "DELETE FROM c_02 WHERE [num_interno] = " & nicgr
   cn1.Execute QUERY
    
   QUERY = "DELETE FROM c_03 WHERE [num_interno] = " & nicgr
   cn1.Execute QUERY
   cn1.CommitTrans
   
   If t_entro = "S" Then
     'emito asiento
     Call graboasiento
   End If
   
End If
Else
 MsgBox ("Periodo cerrado. Imposible realizar operacion")
End If
Me.Hide
End Sub
Sub graboasiento()
      numintcgr = saca_ultnumero_int_comp("G")
      q = "select * from cyb_01 where [id_forma_pago] = " & Val(t_idbanco)
      Set rs = New ADODB.Recordset
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
         
         If t_tipomov = 1 Then
           cta = rs("id_cuenta_ch_dif")
           u1 = "D"
           u2 = "H"
         Else
           cta = rs("id_cuenta_dep_dif")
           u1 = "H"
           u2 = "D"
         
         End If
         
         ctab = rs("id_cuenta_cont")
         
         Set rs1 = New ADODB.Recordset
         q = "select * from c_01 where [id_cuenta] = " & cta
         rs1.Open q, cn1
         If Not rs1.EOF And Not rs1.BOF Then
           dcta = rs("descripcion")
         Else
           dcta = "Cuenta Inexistente"
         End If
         Set rs1 = Nothing
         
         
         Set rs1 = New ADODB.Recordset
         q = "select * from c_01 where [id_cuenta] = " & ctab
         rs1.Open q, cn1
         If Not rs1.EOF And Not rs1.BOF Then
           dctab = rs("descripcion")
         Else
           dctab = "Cuenta Inexistente"
         End If
         Set rs1 = Nothing
         
         cn1.BeginTrans
         'grabo asiento
         QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Conciliacion] Nro. Int." & Format$(Val(t_id), "00000000") & "', 'B', " & Val(t_id) & ", " & Val(t_mov) & ", " & Val(t_mov) & ", " & para.id_usuario & ", 'Conciliacion Bancaria')"
         cn1.Execute QUERY
      
         
         'cuenta madre banco
         ic = 1
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & 1 & ", " & cta & ", '" & u1 & "', " & Val(t_mov) & ", '" & dcta & "')"
         
         cn1.Execute QUERY
         ic = ic + 1
      
         
         'contrapartida
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & ctab & ", '" & u2 & "', " & Val(t_mov) & ", '" & dctab & "')"
         cn1.Execute QUERY
         cn1.CommitTrans

   End If
End Sub

Private Sub btnsale_Click()
Me.Hide
End Sub

Private Sub Form_Activate()
t_entro.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  Call TabEnter2(Me, 1)
End If
End Sub

Private Sub t_entro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  t_fecha.SetFocus
End If

If KeyAscii = 32 Then
  If t_entro = "S" Then
     t_entro = "N"
  Else
     t_entro = "S"
  End If
End If

End Sub

Private Sub t_entro_LostFocus()
t_entro = Format$(t_entro, ">@")
If t_entro <> "S" And t_entro <> "N" Then
  t_entro = "N"
End If
End Sub

Private Sub t_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
   t_entro.SetFocus
End If
End Sub

Private Sub t_fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 btnacepta.SetFocus
End If
End Sub


Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
 If Not IsDate(t_fecha) Then
    t_fecha = Format$(Now, "dd/mm/yyyy")
 End If
Else
 t_fecha = Format$(Now, "dd/mm/yyyy")
End If
 
End Sub


Private Sub UpDown1_DownClick()
t_entro = "S"
End Sub

Private Sub UpDown1_UpClick()
t_entro = "N"
End Sub

Private Sub UpDown2_DownClick()
t_fecha = Format$(DateValue(t_fecha) - 1, "dd/mm/yyyy")

End Sub

Private Sub UpDown2_UpClick()
 t_fecha = Format$(DateValue(t_fecha) + 1, "dd/mm/yyyy")

End Sub
