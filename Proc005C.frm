VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form op_fp2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CHEQUES PROPIOS"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2175
   ScaleWidth      =   13455
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_modulo 
      Height          =   285
      Left            =   9840
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1575
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   13095
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11400
         MaxLength       =   10
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox t_importe 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9000
         MaxLength       =   21
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox t_numch 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6240
         MaxLength       =   9
         TabIndex        =   1
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox c_fp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   0
         Text            =   "c_prod"
         Top             =   840
         Width           =   4695
      End
      Begin VB.TextBox t_fp 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   360
         MaxLength       =   8
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "Fecha Dif."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   11400
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "Importe"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9000
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "Num.Cheque"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6240
         TabIndex        =   8
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "Banco"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   5655
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   13455
      _ExtentX        =   23733
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
            TextSave        =   "29/02/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "04:37 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "op_fp2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   
     
         
End Select
End Sub

Sub modifrenglon()
  If c_fp.ListIndex < 0 Then
    ip = "050"
  Else
    ip = Format$(c_fp.ItemData(c_fp.ListIndex), "000")
  End If
  d = c_fp
  i = Format$(Val(t_importe), "######0.00")
  
  Set rs2 = New ADODB.Recordset
  q = "SELECT * FROM CYB_01 WHERE [ID_FORMA_PAGO] = " & Val(ip)
  rs2.Open q, cn1
  If Not rs2.EOF And Not rs2.BOF Then
   c = rs2("ID_CUENTA_ch_dif")
  Else
   c = 0
  End If
  Set rs2 = Nothing
  
  
  Select Case t_modulo
  Case Is = "O"
    If t_fp <> "" Then
     r = Val(t_fp)
     op.msf2.AddItem ip & Chr(9) & "Ch.Propio" & Chr(9) & Format$(T_NUMCH, "0000000000") & Chr(9) & c_fp & Chr(9) & "" & Chr(9) & glo.nombrecli & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fecha & Chr(9) & "" & Chr(9) & c, r
     op.msf2.RemoveItem r + 1
    Else
     'r = op.msf2.Rows
     op.msf2.AddItem ip & Chr(9) & "Ch.Propio" & Chr(9) & Format$(T_NUMCH, "0000000000") & Chr(9) & c_fp & Chr(9) & "" & Chr(9) & glo.nombrecli & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fecha & Chr(9) & "" & Chr(9) & c
    End If
  Case Is = "D"
    If t_fp <> "" Then
     r = Val(t_fp)
     cyb_depositoS.msf2.AddItem ip & Chr(9) & "Ch.Propio" & Chr(9) & Format$(T_NUMCH, "0000000000") & Chr(9) & c_fp & Chr(9) & "" & Chr(9) & glo.nombrecli & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fecha & Chr(9) & "" & Chr(9) & c, r
     cyb_depositoS.msf2.RemoveItem r + 1
    Else
     'r = op.msf2.Rows
     cyb_depositoS.msf2.AddItem ip & Chr(9) & "Ch.Propio" & Chr(9) & Format$(T_NUMCH, "0000000000") & Chr(9) & c_fp & Chr(9) & "" & Chr(9) & glo.nombrecli & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fecha & Chr(9) & "" & Chr(9) & c
    End If
  
  Case Is = "V"
    If t_fp <> "" Then
     r = Val(t_fp)
     cyb_VENTACH.msf2.AddItem ip & Chr(9) & "Ch.Propio" & Chr(9) & Format$(T_NUMCH, "0000000000") & Chr(9) & c_fp & Chr(9) & "" & Chr(9) & glo.nombrecli & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fecha & Chr(9) & "" & Chr(9) & c, r
     cyb_VENTACH.msf2.RemoveItem r + 1
    Else
     'r = op.msf2.Rows
     cyb_VENTACH.msf2.AddItem ip & Chr(9) & "Ch.Propio" & Chr(9) & Format$(T_NUMCH, "0000000000") & Chr(9) & c_fp & Chr(9) & "" & Chr(9) & glo.nombrecli & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fecha & Chr(9) & "" & Chr(9) & c
    End If
    
   Case Is = "C"
    If t_fp <> "" Then
     r = Val(t_fp)
     com_formapago.msf2.AddItem ip & Chr(9) & "Ch.Propio" & Chr(9) & Format$(T_NUMCH, "0000000000") & Chr(9) & c_fp & Chr(9) & "" & Chr(9) & glo.nombrecli & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fecha & Chr(9) & "" & Chr(9) & c, r
     com_formapago.msf2.RemoveItem r + 1
    Else
     'r = op.msf2.Rows
     com_formapago.msf2.AddItem ip & Chr(9) & "Ch.Propio" & Chr(9) & Format$(T_NUMCH, "0000000000") & Chr(9) & c_fp & Chr(9) & "" & Chr(9) & glo.nombrecli & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fecha & Chr(9) & "" & Chr(9) & c
    End If
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 3)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call carga_formas_pago(c_fp, "B")
c_fp.ListIndex = 0

End Sub

  
Sub limpia()
t_fp = ""
t_importe = ""
T_NUMCH = ""
t_fecha = ""
c_fp.ListIndex = 0
c_fp.SetFocus
End Sub

Private Sub t_fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Not IsDate(t_fecha) Then
   t_fecha = Format$(Now, "dd/mm/yyyy")
 End If
 
 Set rs = New ADODB.Recordset
 q = "select * from cyb_02 where [id_banco] = " & c_fp.ItemData(c_fp.ListIndex) & " and [num_cheque] = " & Val(T_NUMCH)
 rs.Open q, cn1, adOpenDynamic, adLockOptimistic
 If Not rs.EOF And Not rs.BOF Then
   If rs("estado") <> "P" Then
      MsgBox ("El cheque ingresado no se encuentra pendiente para su emision")
   Else
         
      Call modifrenglon
      Call limpia
      Me.Hide
   End If
 Else
   J = MsgBox("El cheque no se encuentra generado. Desea Ingresarlo?", 4)
   If J = 6 Then
      'verificar permisos
      
      rs.AddNew
      rs("id_banco") = c_fp.ItemData(c_fp.ListIndex)
      rs("num_cheque") = Val(T_NUMCH)
      rs("fecha_emision") = Format$(Now, "dd/mm/yyyy")
      rs("fecha_dif") = Format$(Now, "dd/mm/yyyy")
      rs("estado") = "P"
      rs("destino") = "Pendiente"
      rs("importe") = 0
      rs("num_mov_banco") = 0
      rs("id_chequera") = 0
      rs.Update
      Call modifrenglon
      Call limpia
      Me.Hide
    End If
   End If
End If
Set rs = Nothing


End Sub
