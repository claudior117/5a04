VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form cgr_abmasientos_p2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INGRESO DE ITEMS EN ASIENTOS"
   ClientHeight    =   1980
   ClientLeft      =   285
   ClientTop       =   4080
   ClientWidth     =   11190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   11190
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   10680
      TabIndex        =   15
      Top             =   0
      Width           =   375
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   0
         Picture         =   "CGR015B.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Salir sin Modificar"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   375
      End
   End
   Begin VB.TextBox t_ubicaanterior 
      Height          =   285
      Left            =   5280
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   10935
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   5040
         Picture         =   "CGR015B.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   840
         Width           =   375
      End
      Begin VB.ComboBox c_ubica 
         Height          =   315
         ItemData        =   "CGR015B.frx":0B8C
         Left            =   9240
         List            =   "CGR015B.frx":0B96
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox t_importe 
         Height          =   285
         Left            =   8040
         MaxLength       =   12
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox t_detalle 
         Height          =   285
         Left            =   5520
         MaxLength       =   40
         TabIndex        =   2
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox c_cuenta 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox t_cod 
         Height          =   285
         Left            =   720
         MaxLength       =   8
         TabIndex        =   0
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox t_renglon 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   7
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Ubicacion"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9240
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Importe"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7920
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5400
         TabIndex        =   11
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   " Cuenta"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Cod."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1725
      Width           =   11190
      _ExtentX        =   19738
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
            TextSave        =   "31/03/2012"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "06:53 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cgr_abmasientos_p2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984






Private Sub btnsale_Click()
Me.Hide
End Sub

Private Sub c_cuenta_LostFocus()
If c_cuenta.ListIndex < 0 Then
  If Val(c_cuenta) > 0 Then
    c_cuenta.ListIndex = buscaindice(c_cuenta, Val(c_cuenta))
  Else
    c_cuenta.ListIndex = 0
  End If
End If
t_cod = c_cuenta.ItemData(c_cuenta.ListIndex)

End Sub

Private Sub c_ubica_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And c_ubica.ListIndex >= 0 Then
  p = 0
  If Val(t_cod) <= 0 Then
    p = 1
  End If
  If Val(t_importe) <= 0 Then
    p = 1
  End If
  If p = 1 Then
    MsgBox ("Error en los datos Ingresados")
  Else
    If Val(t_renglon) = 0 Then
        'agrego nuevo

      If c_ubica.ListIndex = 0 Then
       'debe
       cgr_abmasientos_p.msf1.AddItem cgr_abmasientos_p.msf1.Rows & Chr$(9) & t_cod & Chr$(9) & t_detalle & Chr$(9) & Format$(Val(t_importe), "######0.00") & Chr$(9) & c_cuenta
      Else
       'haber
       cgr_abmasientos_p.msf2.AddItem cgr_abmasientos_p.msf2.Rows & Chr$(9) & t_cod & Chr$(9) & t_detalle & Chr$(9) & Format$(Val(t_importe), "######0.00") & Chr$(9) & c_cuenta
      End If
    Else
       'modifico
        If t_ubicaanterior = "D" Then
           If cgr_abmasientos_p.msf1.Rows > 2 Then
               cgr_abmasientos_p.msf1.RemoveItem (Val(t_renglon))
           Else
               cgr_abmasientos_p.armagrid
           End If
           If c_ubica.ListIndex = 0 Then
              'modifico el mismo renglon
               cgr_abmasientos_p.msf1.AddItem t_renglon & Chr$(9) & t_cod & Chr$(9) & t_detalle & Chr$(9) & Format$(Val(t_importe), "######0.00") & Chr$(9) & c_cuenta, Val(t_renglon)
           Else
              'saco de debe y paso a haber
               cgr_abmasientos_p.msf2.AddItem cgr_abmasientos_p.msf2.Rows & Chr$(9) & t_cod & Chr$(9) & t_detalle & Chr$(9) & Format$(Val(t_importe), "######0.00") & Chr$(9) & c_cuenta
           End If
       Else
           If cgr_abmasientos_p.msf2.Rows > 2 Then
               cgr_abmasientos_p.msf2.RemoveItem (Val(t_renglon))
           Else
               cgr_abmasientos_p.armagrid2
           End If
         
           If c_ubica.ListIndex = 1 Then
              'modifico el mismo renglon
               cgr_abmasientos_p.msf2.AddItem cgr_abmasientos_p.msf2.Rows & Chr$(9) & t_cod & Chr$(9) & t_detalle & Chr$(9) & Format$(Val(t_importe), "######0.00") & Chr$(9) & c_cuenta, Val(t_renglon)
           Else
              'saco de haber y paso a debe
               cgr_abmasientos_p.msf1.AddItem cgr_abmasientos_p.msf1.Rows & Chr$(9) & t_cod & Chr$(9) & t_detalle & Chr$(9) & Format$(Val(t_importe), "######0.00") & Chr$(9) & c_cuenta
           End If
       
        End If
     End If
    
    cgr_abmasientos_p.calcula_totales
    cgr_abmasientos_p.renumera
    Me.Hide
  End If
End If
End Sub

Private Sub Command3_Click()
cgr_buscacuenta.Show
End Sub

Private Sub Form_Activate()
t_cod.SetFocus
If para.cuenta_sel > 0 Then
  t_cod = para.cuenta_sel
  c_cuenta.ListIndex = buscaindice(c_cuenta, para.cuenta_sel)
  para.cuenta_sel = 0
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call TabEnter2(Me, 4)
End If
End Sub

Private Sub Form_Load()
Call barracgr(Me)
Call carga_cuentas_cont(c_cuenta, "C", "D")
End Sub
Sub limpia()
 t_renglon = ""
 t_cod = ""
 t_detalle = ""
 t_importe = ""
 t_ubicaanterior = ""
 't_cod.SetFocus
End Sub

Private Sub t_cod_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[ENTER] Avanza - [F8] Buscaa Cuentas "
End Sub

Private Sub t_cod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF8 Then
  cgr_buscacuenta.Show
End If
End Sub

Private Sub t_cod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Val(t_cod) > 0 Then
  'busco cuenta
  Set rs = New ADODB.Recordset
  q = "select * from c_01 where [id_cuenta] = " & Val(t_cod)
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
     c_cuenta.ListIndex = buscaindice(c_cuenta, Val(t_cod))
     t_detalle.SetFocus
  Else
     MsgBox ("Cuenta Inexistente")
  End If
 Else
   c_cuenta.SetFocus
 End If
End If
End Sub

Private Sub T_detalle_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[ENTER] Avanza - [F6] Desc. Asiento - [F7] Desc. Cuenta"

End Sub

Private Sub t_detalle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF6 Then
   t_detalle = cgr_abmasientos_p.t_descripciong
End If

If KeyCode = vbKeyF7 Then
   t_detalle = c_cuenta
End If

End Sub

Private Sub t_detalle_LostFocus()
Call barracgr(Me)
End Sub

Private Sub t_importe_KeyPress(KeyAscii As Integer)

Call solonum(KeyAscii, 1)

End Sub
