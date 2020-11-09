VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form gen_descextra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descripcion Extra para Articulos"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox t_funcion 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox t_modulo 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2670
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   123472
            MinWidth        =   123472
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label8 
      Caption         =   "50"
      Height          =   255
      Left            =   6240
      TabIndex        =   11
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "50"
      Height          =   255
      Left            =   6240
      TabIndex        =   10
      Top             =   120
      Width           =   255
   End
   Begin VB.Line Line4 
      X1              =   6120
      X2              =   6120
      Y1              =   120
      Y2              =   2400
   End
   Begin VB.Label Label6 
      Caption         =   "40"
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "40"
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   120
      Width           =   255
   End
   Begin VB.Line Line3 
      X1              =   4920
      X2              =   4920
      Y1              =   120
      Y2              =   2400
   End
   Begin VB.Label Label4 
      Caption         =   "30"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "30"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.Line Line2 
      X1              =   3720
      X2              =   3720
      Y1              =   120
      Y2              =   2400
   End
   Begin VB.Label Label2 
      Caption         =   "20"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "20"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.Line Line1 
      X1              =   2520
      X2              =   2520
      Y1              =   120
      Y2              =   2400
   End
End
Attribute VB_Name = "gen_descextra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()
Text1 = ""
End Sub

Private Sub Command3_Click()
Set rs = New ADODB.Recordset
q = "select * from vta_015"
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  Text1 = rs("desc_ext")
End If
Set rs = Nothing
End Sub

Private Sub Form_Load()
Me.StatusBar1.Panels.Item(1) = "[F9] Agrega - [ESC] Sale sin cambios - [F8] Limpia texto "
Text1 = ""
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF8 Then
  Text1 = ""
End If

If KeyCode = vbKeyF9 Then
 k = SendMessage(Text1.hWnd, EM_GETLINECOUNT, 0, 0&) 'obtiene la cantidad de lieas
 If k <= 5 Then
  
  Select Case t_modulo
  Case Is = "F"
  'factura venta
   If t_funcion = "A" Then
    If Text1 <> "" Then
     vta_facturacion.msf1.AddItem 0 & Chr(9) & "" & Chr(9) & Text1 & Chr(9) & k
      vta_facturacion.msf1.RowHeight(vta_facturacion.msf1.Row + 1) = k * 250
    
    End If
   Else
      vta_facturacion.msf1.AddItem 0 & Chr(9) & "" & Chr(9) & Text1 & Chr(9) & k, vta_facturacion.msf1.Row
      vta_facturacion.msf1.RemoveItem vta_facturacion.msf1.Row + 1
      vta_facturacion.msf1.RowHeight(vta_facturacion.msf1.Row) = k * 250

   End If
   Call vta_facturacion.renumera
  
  Case Is = "E"
  'orden de empaque
   If t_funcion = "A" Then
    If Text1 <> "" Then
     pro_empaque.msf1.AddItem 0 & Chr(9) & "" & Chr(9) & Text1 & Chr(9) & k
     pro_empaque.msf1.RowHeight(pro_empaque.msf1.Row + 1) = k * 250
    
    End If
   Else
      pro_empaque.msf1.AddItem 0 & Chr(9) & "" & Chr(9) & Text1 & Chr(9) & k, pro_empaque.msf1.Row
      pro_empaque.msf1.RemoveItem pro_empaque.msf1.Row + 1
      pro_empaque.msf1.RowHeight(pro_empaque.msf1.Row) = k * 250

   End If
   Call pro_empaque.renumera
  
   Case Is = "P"
  'presdupuesto venta
   If t_funcion = "A" Then
    If Text1 <> "" Then
     vta_presup.msf1.AddItem 0 & Chr(9) & "" & Chr(9) & Text1 & Chr(9) & k
     vta_presup.msf1.RowHeight(vta_presup.msf1.Row + 1) = k * 250
    
    End If
   Else
      vta_presup.msf1.AddItem 0 & Chr(9) & "" & Chr(9) & Text1 & Chr(9) & k, vta_presup.msf1.Row
      vta_presup.msf1.RemoveItem vta_presup.msf1.Row + 1
      vta_presup.msf1.RowHeight(vta_presup.msf1.Row) = k * 250

   End If
   Call vta_presup.renumera
  
  
  Case Is = "O"
  'Orden de Compra
   If t_funcion = "A" Then
    If Text1 <> "" Then
     ABM_OC.msf1.AddItem 0 & Chr(9) & "" & Chr(9) & Text1 & Chr(9) & k, ABM_OC.msf1.Row + 1
     ABM_OC.msf1.RowHeight(ABM_OC.msf1.Row + 1) = k * 250
    
    End If
   Else
      ABM_OC.msf1.AddItem 0 & Chr(9) & "" & Chr(9) & Text1 & Chr(9) & k, ABM_OC.msf1.Row
      ABM_OC.msf1.RemoveItem ABM_OC.msf1.Row + 1
      ABM_OC.msf1.RowHeight(ABM_OC.msf1.Row) = k * 250

   End If
   Call ABM_OC.renumera
  End Select
  Unload Me
  
Else
    MsgBox ("La descripcion extrapara articulos no puede tener mas de 5 lineas")
End If
  
  


End If
  

End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Unload Me
End If

End Sub
