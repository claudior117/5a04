VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_recibo2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INGRESO CHEQUES DE TERCERO"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4500
   ScaleWidth      =   8940
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   4095
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   8415
      Begin VB.TextBox t_modulo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5400
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox t_funcion 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   7680
         MaxLength       =   8
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox t_sucursal 
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   3
         Top             =   2280
         Width           =   6495
      End
      Begin VB.TextBox t_fechad 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox T_NUMCH 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Top             =   840
         Width           =   2055
      End
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
         Left            =   1680
         MaxLength       =   21
         TabIndex        =   5
         Top             =   3240
         Width           =   2175
      End
      Begin VB.TextBox t_titular 
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2760
         Width           =   6495
      End
      Begin VB.TextBox t_banco 
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1800
         Width           =   6495
      End
      Begin VB.TextBox t_NUMINT 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Funcion"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6360
         TabIndex        =   17
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
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Sucursal"
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
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Fecha Dif."
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
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1215
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
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Titular"
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
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Banco"
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
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Num.Ch."
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
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   6
      Top             =   4470
      Width           =   8940
      _ExtentX        =   15769
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
            TextSave        =   "08/03/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:38 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_recibo2"
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



Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 5)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
End Sub



 
  
Sub limpia()
T_NUMCH = ""
t_fechai = ""
t_fechad = ""
t_banco = ""
t_sucursal = ""
t_titular = ""
t_importe = ""

End Sub


Private Sub t_banco_LostFocus()
t_banco = RTrim$(t_banco) & " "

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
  Call modificarenglon
  Call limpia
  Me.Hide
 Else
  Call solonum(KeyAscii, 1)
 End If
End Sub

Sub modificarenglon()
  ip = Format$(3, "000")
  d = "Ch.Terc."
  i = Format$(Val(t_importe), "######0.00")
  Select Case t_modulo
  Case Is = "R"
   If t_fp <> "" Then
     r = Val(t_fp)
     vta_recibo.msf2.AddItem ip & Chr(9) & d & Chr(9) & Format$(T_NUMCH, "0000000000") & Chr(9) & Left$(t_banco, 49) & " " & Chr(9) & Left$(t_sucursal, 49) & " " & Chr(9) & Left$(t_titular, 49) & " " & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fechad & Chr(9) & Chr$(9) & para.cuenta_valores_terceros, r
     vta_recibo.msf2.RemoveItem r + 1
   Else
     vta_recibo.msf2.AddItem ip & Chr(9) & d & Chr(9) & Format$(T_NUMCH, "0000000000") & Chr(9) & Left$(t_banco, 49) & " " & Chr(9) & Left$(t_sucursal, 49) & " " & Chr(9) & Left$(t_titular, 49) & " " & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fechad & Chr(9) & Chr$(9) & para.cuenta_valores_terceros
   End If
  Case Is = "F"
   If t_fp <> "" Then
     r = Val(t_fp)
     vta_formapago.msf2.AddItem ip & Chr(9) & d & Chr(9) & Format$(T_NUMCH, "0000000000") & Chr(9) & Left$(t_banco, 49) & " " & Chr(9) & Left$(t_sucursal, 49) & " " & Chr(9) & Left$(t_titular, 49) & " " & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fechad & Chr$(9) & Chr(9) & para.cuenta_valores_terceros, r
     vta_formapago.msf2.RemoveItem r + 1
   Else
     vta_formapago.msf2.AddItem ip & Chr(9) & d & Chr(9) & Format$(T_NUMCH, "0000000000") & Chr(9) & Left$(t_banco, 49) & " " & Chr(9) & Left$(t_sucursal, 49) & " " & Chr(9) & Left$(t_titular, 49) & " " & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fechad & Chr$(9) & Chr(9) & para.cuenta_valores_terceros
   End If
  
 Case Is = "Q"
   If t_fp <> "" Then
     r = Val(t_fp)
     fsc_formapago.msf2.AddItem ip & Chr(9) & d & Chr(9) & Format$(T_NUMCH, "0000000000") & Chr(9) & Left$(t_banco, 49) & " " & Chr(9) & Left$(t_sucursal, 49) & " " & Chr(9) & Left$(t_titular, 49) & " " & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fechad & Chr$(9) & Chr(9) & para.cuenta_valores_terceros & Chr(9) & "Ch. " & Format$(T_NUMCH, "0000000000") & 3, r
     fsc_formapago.msf2.RemoveItem r + 1
   Else
     fsc_formapago.msf2.AddItem ip & Chr(9) & d & Chr(9) & Format$(T_NUMCH, "0000000000") & Chr(9) & Left$(t_banco, 49) & " " & Chr(9) & Left$(t_sucursal, 49) & " " & Chr(9) & Left$(t_titular, 49) & " " & Chr(9) & Format$(t_importe, "######0.00") & Chr(9) & t_fechad & Chr$(9) & Chr(9) & para.cuenta_valores_terceros & Chr(9) & "Ch. " & Format$(T_NUMCH, "0000000000") & 3
   End If
  
  End Select

End Sub

Private Sub T_NUMCH_LostFocus()
T_NUMCH = Format$(Val(T_NUMCH), "00000000")
End Sub

Private Sub t_sucursal_LostFocus()
t_sucursal = RTrim$(t_sucursal) & " "

End Sub

Private Sub t_titular_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF6 Then
  t_titular = vta_recibo.denominACION
End If
End Sub

Private Sub t_titular_LostFocus()
t_titular = RTrim$(t_titular) & " "

End Sub
