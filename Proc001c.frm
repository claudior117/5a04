VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form abm_oc1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2175
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1575
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   11295
      Begin VB.TextBox t_nroreq 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   10080
         MaxLength       =   8
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox t_cantunit 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   9120
         MaxLength       =   8
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox t_envase 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   8160
         MaxLength       =   8
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox c_prod 
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
         Left            =   1320
         TabIndex        =   0
         Text            =   "c_prod"
         Top             =   840
         Width           =   5655
      End
      Begin VB.TextBox t_cantidad 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   7200
         MaxLength       =   8
         TabIndex        =   1
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox t_renglon 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   360
         MaxLength       =   8
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Nro. Requisicion"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10080
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Cantidad Unitaria"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9120
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Envase"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8160
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Cantidad a Pedir"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7080
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Producto"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   6735
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   11910
      _ExtentX        =   21008
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
            TextSave        =   "22/12/2005"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:07 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "abm_oc1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   
     
         
End Select
End Sub

Sub modifrenglon()
If Val(t_cantidad) >= 0 Then
    If c_prod.ListIndex <= 0 Then
    ip = 1
  Else
    ip = c_prod.ItemData(c_prod.ListIndex)
  End If
  d = c_prod
  c = Format$(Val(t_cantidad), "######0.00")
  e = Format$(Val(t_envase), "######0.00")
  cu = Format$(Val(t_cantunit), "######0.00")
  nr = Format$(Val(t_nroreq), "00000")
  If t_renglon <> "" Then
     r = t_renglon
     ABM_OC.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & "0.00" & Chr(9) & cu & Chr(9) & e & Chr(9) & c & Chr(9) & nr, Val(t_renglon)
     ABM_OC.msf1.RemoveItem Val(t_renglon) + 1
  Else
     r = ABM_OC.msf1.Rows
     ABM_OC.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & "0.00" & Chr(9) & cu & Chr(9) & e & Chr(9) & c & Chr(9) & nr
  End If
End If

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
Call carga_productos(c_prod)
c_prod.ListIndex = 0

End Sub


Private Sub t_cantidad_KeyPress(KeyAscii As Integer)
  Call solonum(KeyAscii, 1)

End Sub

 
  
Sub limpia()
t_cantidad = ""
t_renglon = ""
t_envase = ""
t_cantunit = ""
c_prod.ListIndex = 0
c_prod.SetFocus
End Sub

Private Sub t_cantunit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call modifrenglon
  Call limpia
  Me.Hide
Else
   Call solonum(KeyAscii, 1)
End If

End Sub

Private Sub t_envase_LostFocus()
t_cantunit = Format$(Val(t_cantidad) * Val(t_envase), "#####0.00")
End Sub
