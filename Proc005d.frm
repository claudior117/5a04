VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form op_fp1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CARTERA DE CHEQUES DE TERCERO"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6600
   ScaleWidth      =   10545
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Orden"
      Height          =   615
      Left            =   6240
      TabIndex        =   9
      Top             =   5640
      Width           =   3015
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Dif."
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Num.Int."
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Seleccion"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   5895
      Begin VB.TextBox t_importe 
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_cantidad 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Importe:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Cantidad:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comprobantes a Aplicar"
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10335
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   5175
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   9128
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox t_modulo 
      Height          =   285
      Left            =   9240
      MaxLength       =   1
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6345
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   13229
            MinWidth        =   13229
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "27/02/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:44"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "op_fp1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Private Sub Form_Activate()
Call cargacarterach
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 7
msf1.ColWidth(0) = 300
msf1.ColWidth(1) = 800
msf1.ColWidth(2) = 1000
msf1.ColWidth(3) = 1200
msf1.ColWidth(4) = 3000
msf1.ColWidth(5) = 2000
msf1.ColWidth(6) = 1200

msf1.TextMatrix(0, 0) = ""
msf1.TextMatrix(0, 1) = "Num.Int."
msf1.TextMatrix(0, 2) = "Num.Ch."
msf1.TextMatrix(0, 3) = "Fecha Dif."
msf1.TextMatrix(0, 4) = "Banco"
msf1.TextMatrix(0, 5) = "Sucursal"
msf1.TextMatrix(0, 6) = "Importe"
t_importe = ""
t_cantidad = ""

End Sub

Sub carga()
r = 1
Set rs = New ADODB.Recordset
q = "SELECT * FROM CYB_01 WHERE [ID_FORMA_PAGO] = 3"
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
   c = rs("ID_CUENTA_CONT")
Else
   c = 0
End If
Set rs = Nothing

If msf1.Rows > 1 Then
 For i = 1 To msf1.Rows - 1
  If msf1.TextMatrix(i, 0) = "**" Then
    ni = msf1.TextMatrix(i, 1)
   nch = msf1.TextMatrix(i, 2)
   fd = msf1.TextMatrix(i, 3)
   b = msf1.TextMatrix(i, 4)
   t = msf1.TextMatrix(i, 5)
   im = msf1.TextMatrix(i, 6)
   Select Case t_modulo
     Case Is = "O" 'ordenes de pago
       op.msf2.AddItem "003" & Chr(9) & "Ch.Terc." & Chr(9) & nch & Chr(9) & b & Chr(9) & " " & Chr(9) & t & Chr(9) & im & Chr(9) & fd & Chr(9) & ni & Chr(9) & c
     Case Is = "D" 'depositos
       cyb_depositoS.msf2.AddItem "003" & Chr(9) & "Ch.Terc." & Chr(9) & nch & Chr(9) & b & Chr(9) & " " & Chr(9) & t & Chr(9) & im & Chr(9) & fd & Chr(9) & ni & Chr(9) & c
     Case Is = "V" 'venta ch.
       cyb_VENTACH.msf2.AddItem "003" & Chr(9) & "Ch.Terc." & Chr(9) & nch & Chr(9) & b & Chr(9) & " " & Chr(9) & t & Chr(9) & im & Chr(9) & fd & Chr(9) & ni & Chr(9) & c
     Case Is = "C" 'comprobantes contado
       com_formapago.msf2.AddItem "003" & Chr(9) & "Ch.Terc." & Chr(9) & nch & Chr(9) & b & Chr(9) & " " & Chr(9) & t & Chr(9) & im & Chr(9) & fd & Chr(9) & ni & Chr(9) & c
   
   End Select
  End If
 Next i
End If
   
End Sub


Private Sub cargacarterach()
Call armagrid
Set rs = New ADODB.Recordset
q = "select * from cyb_03 where [estado] = " & "'" & "C" & "'"
If Option1 = True Then
  q = q & " order by [num_interno]"
Else
  q = q & " order by [fecha_dif], [num_interno]"
End If
rs.Open q, cn1
tot = 0
c = 0
While Not rs.EOF
     msf1.AddItem "" & Chr$(9) & Format$(rs("num_interno"), "00000") & Chr$(9) & Format$(rs("num_cheque"), "0000000000") & Chr$(9) & Format$(rs("fecha_dif"), "dd/mm/yyyy") & Chr$(9) & Format$(Left$(rs("banco"), 25), ">@@@@@@@@@@@@@@@@@@@@@@@@@!") & Chr$(9) & Format$(Left$(rs("titular"), 25), ">@@@@@@@@@@@@@@@@@@@@@@@@@!") & Chr$(9) & Format$(rs("importe"), "######0.00")
     tot = tot + rs("importe")
     c = c + 1
     rs.MoveNext
Wend
Set rs = Nothing
msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "_________________________"
msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "Cant. Cheques: " & c & Chr$(9) & "" & Chr$(9) & Format$(tot, "######0.00")
t_cantidad = "0"
t_importe = "0.00"

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Unload Me
End If
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Option1 = True
End Sub

  






Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
  If msf1.Rows > 1 Then
    t_cantidad = "0"
    t_importe = "0.00"
    For i = 1 To msf1.Rows - 1
      If msf1.TextMatrix(i, 0) = "**" Then
          msf1.TextMatrix(i, 0) = ""
          t_cantidad = Val(t_cantidad) - 1
          t_importe = Format$(Val(t_importe) - Val(msf1.TextMatrix(i, 6)), "#####0.00")
      Else
         msf1.TextMatrix(i, 0) = "**"
          t_cantidad = Val(t_cantidad) + 1
          t_importe = Format$(Val(t_importe) + Val(msf1.TextMatrix(i, 6)), "#####0.00")
      End If
    Next i
  End If
End If


If KeyCode = vbKeyInsert Then
  op_fp1_1.Show
  op_fp1_1.t_funcion = "A"
  
End If

If KeyCode = vbKeyF9 Then
  Call carga
  Me.Hide
End If

End Sub


Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[Barra] Selecciona - [F5] Todos  - [F9] Agrega - [INS] Nuevo CH. - [ENTER] Modif. Ch. "
msf1.FocusRect = flexFocusNone
End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeySpace Then
  If Val(msf1.TextMatrix(msf1.Row, 1)) > 0 Then
      If msf1.TextMatrix(msf1.Row, 0) = "**" Then
          msf1.TextMatrix(msf1.Row, 0) = ""
          t_cantidad = Val(t_cantidad) - 1
          t_importe = Format$(Val(t_importe) - Val(msf1.TextMatrix(msf1.Row, 6)), "#####0.00")

      Else
         msf1.TextMatrix(msf1.Row, 0) = "**"
         t_cantidad = Val(t_cantidad) + 1
         t_importe = Format$(Val(t_importe) + Val(msf1.TextMatrix(msf1.Row, 6)), "#####0.00")
      
      End If
  End If
  
End If

End Sub

Private Sub Option1_Click()
Call cargacarterach
End Sub

Private Sub Option2_Click()
Call cargacarterach
End Sub
