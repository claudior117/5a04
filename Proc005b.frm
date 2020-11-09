VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form op_fp3 
   BackColor       =   &H00C0C0C0&
   Caption         =   "FORMA DE PAGO"
   ClientHeight    =   2175
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2175
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_modulo 
      Height          =   285
      Left            =   8640
      MaxLength       =   1
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1575
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   11775
      Begin VB.ComboBox c_cuenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8280
         TabIndex        =   3
         Text            =   "c_prod"
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox t_importe 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   6960
         MaxLength       =   10
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox t_detalle 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   3120
         MaxLength       =   25
         TabIndex        =   1
         Top             =   840
         Width           =   3735
      End
      Begin VB.ComboBox c_fp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1080
         TabIndex        =   0
         Text            =   "c_prod"
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox t_fp 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   120
         MaxLength       =   8
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Cuenta"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8280
         TabIndex        =   11
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Importe"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6960
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Forma Pago"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3015
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1920
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
            TextSave        =   "13/10/2010"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "07:14 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "op_fp3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984


Private Sub c_cuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If c_cuenta.ListIndex >= 0 Then
    Call modifrenglon
    Call limpia
    Me.Hide
  End If
End If

End Sub

Private Sub c_fp_LostFocus()
If c_fp.ListIndex < 0 Then
  c_fp.ListIndex = 0
End If
Set rs = New ADODB.Recordset
q = "select * from cyb_01 where [id_forma_pago] = " & c_fp.ItemData(c_fp.ListIndex)
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
  cuenta = rs("id_cuenta_cont")
Else
  cuenta = 0
End If
Set rs = Nothing
c_cuenta.ListIndex = buscaindice(c_cuenta, cuenta)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   
     
         
End Select
End Sub

Sub modifrenglon()
  If c_fp.ListIndex <= 0 Then
    ip = "001"
  Else
    ip = Format$(c_fp.ItemData(c_fp.ListIndex), "000")
  End If
  d = c_fp
  i = Format$(Val(t_importe), "######0.00")
  c = c_cuenta.ItemData(c_cuenta.ListIndex)
  Select Case t_modulo
  Case Is = "O" 'op
    If t_fp <> "" Then
     r = Val(t_fp)
     op.msf2.AddItem ip & Chr(9) & c_fp & Chr(9) & "-" & Chr(9) & t_detalle & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & Format$(Now, "DD/MM/YYYY") & Chr(9) & "" & Chr(9) & c, r
     op.msf2.RemoveItem r + 1
    Else
     'r = op.msf2.Rows
     op.msf2.AddItem ip & Chr(9) & c_fp & Chr(9) & "-" & Chr(9) & t_detalle & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & Format$(Now, "DD/MM/YYYY") & Chr(9) & "" & Chr(9) & c
    End If
  Case Is = "D" 'op
    If t_fp <> "" Then
     r = Val(t_fp)
     cyb_depositoS.msf2.AddItem ip & Chr(9) & c_fp & Chr(9) & "-" & Chr(9) & t_detalle & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & Format$(Now, "DD/MM/YYYY") & Chr(9) & "" & Chr(9) & c, r
     cyb_depositoS.msf2.RemoveItem r + 1
    Else
     'r = op.msf2.Rows
     cyb_depositoS.msf2.AddItem ip & Chr(9) & c_fp & Chr(9) & "-" & Chr(9) & t_detalle & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & Format$(Now, "DD/MM/YYYY") & Chr(9) & "" & Chr(9) & c
    End If
  Case Is = "F" 'fact. compra contado
    If t_fp <> "" Then
     r = Val(t_fp)
     com_formapago.msf2.AddItem ip & Chr(9) & c_fp & Chr(9) & "-" & Chr(9) & t_detalle & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & Format$(Now, "DD/MM/YYYY") & Chr(9) & "" & Chr(9) & c, r
     com_formapago.msf2.RemoveItem r + 1
    Else
     com_formapago.msf2.AddItem ip & Chr(9) & c_fp & Chr(9) & "-" & Chr(9) & t_detalle & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & Format$(Now, "DD/MM/YYYY") & Chr(9) & "" & Chr(9) & c
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
Call carga_formas_pago(c_fp, "O")
c_fp.ListIndex = 0

Call carga_cuentas_cont(c_cuenta, "C", "D")
c_cuenta.AddItem "<Sin Imputacion>", 0
c_cuenta.ListIndex = 0
End Sub

  
Sub limpia()
t_fp = ""
t_importe = ""
t_detalle = ""
c_fp.ListIndex = 0
c_fp.SetFocus
End Sub

Private Sub t_importe_KeyPress(KeyAscii As Integer)
   Call solonum(KeyAscii, 1)

End Sub

