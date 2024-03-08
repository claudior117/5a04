VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_recibo3 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORMA DE PAGO"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12570
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2175
   ScaleWidth      =   12570
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_modulo 
      Height          =   285
      Left            =   11280
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   12135
      Begin VB.TextBox t_importe 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9840
         MaxLength       =   21
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox t_detalle 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   1
         Top             =   720
         Width           =   4935
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
         Left            =   1440
         TabIndex        =   0
         Text            =   "c_prod"
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox t_fp 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   360
         MaxLength       =   8
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9840
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4800
         TabIndex        =   7
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Forma Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   4575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   1770
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   21167
            MinWidth        =   21167
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_recibo3"
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
  If c_fp.ListIndex <= 0 Then
    ip = "001"
  Else
    ip = Format$(c_fp.ItemData(c_fp.ListIndex), "000")
  End If
  Set rs = New ADODB.Recordset
  q = "select * from cyb_01  where [id_forma_pago] = " & Val(ip)
  rs.MaxRecords = 1
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
    cta = rs("id_cuenta_cont")
  Else
    cta = 0
  End If
  Set rs = Nothing
  d = c_fp
  i = Format$(Val(t_importe), "######0.00")
  
  
  Select Case t_modulo
  Case Is = "R"
   If t_fp <> "" Then
     r = Val(t_fp)
     vta_recibo.msf2.AddItem ip & Chr(9) & c_fp & Chr(9) & "-" & Chr(9) & t_detalle & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & Format$(Now, "DD/MM/YYYY") & Chr(9) & Chr(9) & cta, r
     vta_recibo.msf2.RemoveItem r + 1
   Else
     vta_recibo.msf2.AddItem ip & Chr(9) & c_fp & Chr(9) & "-" & Chr(9) & t_detalle & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & Format$(Now, "DD/MM/YYYY") & Chr(9) & Chr(9) & cta
   End If
  Case Is = "F"
    If t_fp <> "" Then
     r = Val(t_fp)
     vta_formapago.msf2.AddItem ip & Chr(9) & c_fp & Chr(9) & "-" & Chr(9) & t_detalle & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & Format$(vta_facturacion.t_fecha, "DD/MM/YYYY") & Chr$(9) & Chr(9) & cta, r
     vta_formapago.msf2.RemoveItem r + 1
   Else
     vta_formapago.msf2.AddItem ip & Chr(9) & c_fp & Chr(9) & "-" & Chr(9) & t_detalle & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & Format$(vta_facturacion.t_fecha, "DD/MM/YYYY") & Chr$(9) & Chr(9) & cta
   End If
   Case Is = "Q"
    If Len(t_detalle) <= 2 Then
      t_detalle = c_prod
    End If
    If t_fp <> "" Then
     r = Val(t_fp)
     fsc_formapago.msf2.AddItem ip & Chr(9) & c_fp & Chr(9) & "-" & Chr(9) & t_detalle & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & Format$(fsc_tique.t_fecha, "DD/MM/YYYY") & Chr$(9) & Chr(9) & cta & Chr(9) & Left$(t_detalle, 15), r
     fsc_formapago.msf2.RemoveItem r + 1
   Else
     fsc_formapago.msf2.AddItem ip & Chr(9) & c_fp & Chr(9) & "-" & Chr(9) & t_detalle & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & Format$(fsc_tique.t_fecha, "DD/MM/YYYY") & Chr$(9) & Chr(9) & cta & Chr(9) & Left$(t_detalle, 15)
   End If
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
Call carga_formas_pago(c_fp, "Y")
c_fp.ListIndex = 0

End Sub

  
Sub limpia()
t_fp = ""
t_importe = ""
t_detalle = ""
c_fp.ListIndex = 0
c_fp.SetFocus
End Sub

Private Sub t_importe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call modifrenglon
  Call limpia
  Me.Hide
Else
   Call solonum(KeyAscii, 1)
End If

End Sub

