VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fsc_facturacion 
   BackColor       =   &H00E0E0E0&
   Caption         =   "FACTURACION FISCAL"
   ClientHeight    =   6780
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   8205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame11 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   240
      TabIndex        =   59
      Top             =   7440
      Width           =   6015
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9000
      TabIndex        =   54
      Top             =   5280
      Width           =   2775
      Begin VB.CommandButton Command4 
         Caption         =   "Cal"
         Height          =   255
         Left            =   2280
         TabIndex        =   56
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Por canje Cereal RG 2459"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Totales"
      Height          =   855
      Left            =   6360
      TabIndex        =   51
      Top             =   7440
      Width           =   2535
      Begin VB.TextBox t_total 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox T_total2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   19
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "Total"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Total U$s"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   52
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Parciales"
      Height          =   855
      Left            =   240
      TabIndex        =   48
      Top             =   6600
      Width           =   4095
      Begin VB.CommandButton Command7 
         Caption         =   "IVA"
         Height          =   195
         Left            =   2760
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox t_iva 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox t_subtotal 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox t_nograbado 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "Subtotal"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "No Grabado"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   49
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Percepciones"
      Height          =   855
      Left            =   4440
      TabIndex        =   44
      Top             =   6600
      Width           =   4455
      Begin VB.CommandButton Command6 
         Caption         =   "Percepcion IB"
         Height          =   195
         Left            =   120
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox t_perciva 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   17
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox t_alicuotaperciva 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   16
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox t_perc 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   15
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox t_alicuotaib 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         MaxLength       =   8
         TabIndex        =   14
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "%"
         Height          =   255
         Left            =   2520
         TabIndex        =   47
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "Perc. Iva"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   46
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Caption         =   "%"
         Height          =   255
         Left            =   600
         TabIndex        =   45
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   " Remitos"
      Height          =   255
      Left            =   8160
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   9720
      TabIndex        =   35
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton Command8 
         Caption         =   "F.P."
         Height          =   255
         Left            =   1080
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contado "
         Height          =   255
         Left            =   120
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cta.Cte."
         Height          =   255
         Left            =   120
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9360
      TabIndex        =   32
      Top             =   2400
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Funcion"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Totales del Comprobante"
      Height          =   1335
      Left            =   240
      TabIndex        =   30
      Top             =   5280
      Width           =   8655
      Begin VB.ComboBox c_actividad 
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   600
         Width           =   5055
      End
      Begin VB.CommandButton Command1 
         Height          =   255
         Left            =   6720
         Picture         =   "fsc_003A.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox t_observaciones 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   10
         Top             =   960
         Width           =   6855
      End
      Begin VB.ComboBox c_vend 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Actividad:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Vendedor"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   3255
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5741
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1935
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   9375
      Begin VB.ComboBox c_sucursal 
         Height          =   315
         ItemData        =   "fsc_003A.frx":0105
         Left            =   7440
         List            =   "fsc_003A.frx":0107
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Height          =   255
         Left            =   8520
         Picture         =   "fsc_003A.frx":0109
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox t_fechavto 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox t_cotizacion 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7080
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Left            =   7680
         Picture         =   "fsc_003A.frx":047B
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox t_letra 
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
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   20
         Top             =   1200
         Width           =   375
      End
      Begin VB.ComboBox c_tipocomp 
         Height          =   315
         ItemData        =   "fsc_003A.frx":0580
         Left            =   1680
         List            =   "fsc_003A.frx":0582
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox t_numcomp 
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
         Height          =   285
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
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
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   2
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Text            =   "c_prov"
         Top             =   720
         Width           =   6015
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Punto Venta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5880
         TabIndex        =   64
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Fecha Vto.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   57
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Cotizacion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   43
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Nro. Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Cliente"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   22
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "fsc_003A.frx":0584
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "fsc_003A.frx":0E06
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   6525
      Width           =   8205
      _ExtentX        =   14473
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
            TextSave        =   "27/08/2009"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:37 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "fsc_facturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Dim cantidadp As Double
Dim calcula_perc_ib As String
Dim alicuota_perc_ib As Single
Dim minimo_perc_ib As Double
Dim gcuit As String
Dim numint As Long

Sub limpia()
   Call armagrid
   t_subtotal = ""
   t_nograbado = ""
   t_perc = ""
   t_iva = ""
   t_total = ""
   Option1 = True
   
End Sub
Sub carga()
  Set rs = New ADODB.Recordset
  q = " select * from vta_02 where [sucursal] = " & Val(t_sucursal) & " and letra = '" & t_letra & "' and [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & " and [num_comp] = " & Val(t_numcomp)
  rs.Open q, cn1
  If Not rs.BOF And Not rs.EOF Then
     MsgBox ("Comprobante Existente")
     EXISTE = "S"
     t_fecha = rs("fecha")
     c_prov.ListIndex = buscaindice(c_prov, rs("id_cliente"))
     
     Set rs1 = New ADODB.Recordset
     q = "select * from vta_03 where [num_int] = " & rs("num_int")
     rs1.Open q, cn1
     Call armagrid
     While Not rs1.EOF
        r = msf1.Rows
        msf1.AddItem r & Chr(9) & Format$(rs1("id_producto"), "00000") & Chr(9) & rs1("descripcion") & Chr(9) & rs1("cantidad") & Chr(9) & rs1("unidad") & Chr$(9) & Format$(rs1("pu"), "######0.00") & Chr(9) & rs1("tasaiva") & Chr(9) & rs1("importe") & Chr(9) & rs1("pu_final")
        rs1.MoveNext
     Wend
     Set rs1 = Nothing
     c_vend.ListIndex = buscaindice(c_vend, rs("id_vendedor"))
     t_subtotal = Format$(rs("subtotal"), "######0.00")
     t_nograbado = Format$(rs("impuestos"), "######0.00")
     t_perc = Format$(rs("perc_iva") + rs("perc_gan") + rs("perc_ib"), "######0.00")
     t_iva = Format$(rs("iva"), "######0.00")
     t_total = Format$(rs("total"), "######0.00")
     
     vta_formapago.armagrid2
     If rs("contado") = "S" Then
        vta_clientes.t_cli = rs("cliente02")
        vta_clientes.t_direccion = rs("direccion02")
        vta_clientes.t_cuit = rs("cuit02")
        vta_clientes.t_localidad = rs("localidad02")
        vta_clientes.c_iva.ListIndex = buscaindice(vta_clientes.c_iva, rs("id_tipo_iva02"))
        Option2 = True
        
        
        
        
        
    End If
     
     Set rs = Nothing
  
  
   'formas de pago
   
  
  
  Else
     EXISTE = "N"
  End If
  
End Sub

Private Sub btnacepta_Click()
If c_prov.ListIndex > 0 Then
  Call iniciagraba
Else
  If Option2 = True Then
    If Val(vta_formapago.t_diferencia) = 0 And vta_formapago.msf2.Rows > 1 And vta_formapago.msf2.Rows <= 6 Then
       Call iniciagraba
    Else
       If vta_formapago.msf2.Rows <= 1 Then
           J = MsgBox("No ha ingresado forma de pago, acepta pago total en Efectivo", 4)
           If J = 6 Then
              'pone forma de pago efectivo
              vta_formapago.msf2.AddItem "001" & Chr(9) & 1 & Chr(9) & "-" & Chr(9) & "Efectivo $" & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & Format$(Val(t_total), "######0.00") & Chr(9) & Format$(t_fecha, "DD/MM/YYYY")
              Call iniciagraba
           End If
       Else
        If Val(vta_formapago.t_diferencia) <> 0 Then
          MsgBox ("El pago ingresado no coincide con el total del comprobante")
        Else
          MsgBox ("En comprobantes fiscales se pueden ingresar hasta 6 pagos solamente")
        End If
       End If
    End If
  Else
    MsgBox ("El Cliente Manual solo puede utilizarse para facturacion de contado")
  End If
End If



End Sub
Sub iniciagraba()
J = MsgBox("Emite Comprobante Fiscal", 4)
If J = 6 Then
  seguir = True
  While seguir
    If imprime_facturafiscal Then
        espere.ProgressBar1.Value = 5
        espere.Label1 = "Espere... Grabando Comprobante Fiscal"
        Call graba
        seguir = False
    Else
        J = MsgBox("Error al Imprimir el Comprobante. Verifique la Impresora para continuar.  Reintente o Cancele", 5)
        If J = 4 Then
           seguir = True
        Else
           seguir = False
        End If
    End If
    Unload espere
  Wend
End If
  
  
  
  
  
End Sub

Function imprime_facturafiscal() As Boolean
Dim CUIT As String
Dim identifica As String
Dim tpago As String
Dim t As String

espere.Show
espere.Refresh
espere.ProgressBar1.Min = 0
espere.ProgressBar1.Max = 6
espere.ProgressBar1.Value = 1
espere.Label1 = "Espere... Comprobando Impresora"
'abrir factura
If vta_clientes.c_iva.ItemData(c_iva.ListIndex) <> 3 Then
   identifica = "CUIT"
   CUIT = Mid$(vta_clientes.t_cuit, 1, 2) & Mid$(vta_clientes.t_cuit, 4, 8) & Mid$(vta_clientes.t_cuit, 13, 1)
 Else
   identifica = "DNI"
   CUIT = "0"
 End If
 
 If Option1 = True Then
    tpago = "Importe en Cta.Cte. Nro. " & Format$(c_prov.ItemData(c_prov.ListIndex), "00000")
 Else
    tpago = "CONTADO "
 End If


 Call NULOS(t_remito)
 espere.ProgressBar1.Value = 2
 espere.Label1 = "Espere... Abriendo Comprobante Fiscal:" & c_tipocomp
 
 t_remito = " "
 
 Select Case c_tipocomp.ItemData(c_tipocomp.ListIndex)
 Case Is = 315 '    'TIQUE factura fiscal
   r = epson1.OpenInvoice("T", "C", t_letra, "1", "P", "17", "I", vta_clientes.t_codfiscal, Left$(vta_clientes.t_cli & "-", 40), " ", identifica, CUIT, "N", Left$(vta_clientes.t_direccion & "-", 40), Left$(vta_clientes.t_localidad, 40), tpago, Left$("Remitos: " & t_remito, 40), " ", "C")
 Case Is = 320 '    'factura fiscal
   r = epson1.OpenInvoice("F", "C", t_letra, "1", "P", "17", "I", vta_clientes.t_codfiscal, Left$(vta_clientes.t_cli & "-", 40), " ", identifica, CUIT, "N", Left$(vta_clientes.t_direccion.t_direccion & "-", 40), Left$(vta_clientes.t_localidad, 40), tpago, Left$("Remitos: " & t_remito, 40), " ", "C")
 Case Is = 321 'nc fiscal
   r = epson1.OpenInvoice("N", "C", t_letra, "1", "P", "17", "I", vta__clientes.t_codfiscal, Left$(vta_clientes.t_cli & "-", 40), " ", identifica, CUIT, "N", Left$(vta_clientes.t_direccion.t_direccion & "-", 40), Left$(vta_clientes.t_localidad, 40), tpago, Left$("Remitos: " & t_remito, 40), " ", "C")
 Case Is = 322       'nd
   r = epson1.OpenInvoice("D", "C", t_letra, "1", "P", "17", "I", vta__clientes.t_codfiscal, Left$(vta_clientes.t_cli & "-", 40), " ", identifica, CUIT, "N", Left$(vta_clientes.t_direccion.t_direccion & "-", 40), Left$(vta_clientes.t_localidad, 40), tpago, Left$("Remitos: " & t_remito, 40), " ", "C")
 End Select
 
 'envia items a facturar
espere.ProgressBar1.Value = 3
espere.Label1 = "Espere... Imprimiendo Productos"
 
 i = 1
 While r And i < msf1.Rows
      If r Then
         r = epson1.SendInvoiceItem(Left$(msf1.TextMatrix(i, 2), 20), Format$(Val(msf1.TextMatrix(i, 3)) * 1000, "00000000"), Format$(Val(msf1.TextMatrix(i, 4)) * 100, "000000000"), Format$(Val(msf1.TextMatrix(i, 5)) * 100, "0000"), "M", "0", "0", " ", " ", " ", "0", "0")
      Else
        i = msf1.Rows
      End If
   i = i + 1
 Wend
 
 
 'pagos
  espere.Label1 = "Espere... Grabando Pagos"
  
  If Option2 = True Then 'contado
   For i = 1 To vta_formapago.msf2.Rows - 1
     td = Left$(RTrim$(vta_formapago.msf2.TextMatrix(i, 2)), 15)
     mp = Format$(Val(vta_formapago.msf2.TextMatrix(i, 3)) * 1000, "00000000")
     dp = "T"
      If r Then
         r = epson1.SendInvoicePayment(td, mp, dp)
      Else
         Call captura
      End If
   Next i
  Else
    td = "Cta. Cte. Nro. " & Format$(c_prov.ItemData(c_prov.ListIndex), "00000")
     mp = Format$(Val(t_total) * 1000, "00000000")
     dp = "T"
     If r Then
       r = epson1.SendInvoicePayment(td, mp, dp)
     Else
       Call captura
      End If
  End If
  
 'subtotal para obtener el importe neto, iva y total impreso en la factura
espere.ProgressBar1.Value = 4
espere.Label1 = "Espere... Cerrando Comprobante Fiscal"

 If r Then
    r = epson1.GetInvoiceSubtotal("N")
 Else
    Call captura
 End If
 If r Then
      t_subtotal = Format$(Val(epson1.AnswerField_10) / 100, "######0.00")
      t_iva = Format$(Val(epson1.AnswerField_6) / 100, "####0.00")
      t_total = Format$(Val(epson1.AnswerField_5) / 100, "######0.00")
 End If
 
 
  Select Case c_tipocomp.ItemData(c_tipocomp.ListIndex)
 Case Is = 315
     If r Then r = epson1.CloseInvoice("T", t_letra, " ")
  Case Is = 320
     If r Then r = epson1.CloseInvoice("F", t_letra, " ")
  Case Is = 321
     If r Then r = epson1.CloseInvoice("N", t_letra, " ")
  Case Is = 322
        If r Then r = epson1.CloseInvoice("D", t_letra, " ")
  End Select
   
  If r Then t_numcomp = epson1.AnswerField_3
   
  imprime_facturafiscal = r
    
   
End Function
Sub captura()
MsgBox ("Estado Fiscal: " & epson1.FiscalStatus & "  Estado Impresora: " & epson1.PrinterStatus)
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 10
msf1.ColWidth(0) = 500
msf1.ColWidth(1) = 700
msf1.ColWidth(2) = 5000
msf1.ColWidth(3) = 1100
msf1.ColWidth(4) = 900
msf1.ColWidth(5) = 1100
msf1.ColWidth(6) = 900
msf1.ColWidth(7) = 1100
msf1.ColWidth(8) = 1100
msf1.ColWidth(9) = 1100
msf1.TextMatrix(0, 0) = "Reng."
msf1.TextMatrix(0, 1) = "Id.Prod."
msf1.TextMatrix(0, 2) = "Detalle"
msf1.TextMatrix(0, 3) = "Cantidad"
msf1.TextMatrix(0, 4) = "Unidad"
msf1.TextMatrix(0, 5) = "P.U."
msf1.TextMatrix(0, 6) = "% Iva"
msf1.TextMatrix(0, 7) = "Importe"
msf1.TextMatrix(0, 8) = "PU Final"
msf1.TextMatrix(0, 9) = "Iva"

End Sub


Private Sub c_actividad_LostFocus()
If c_actividad.ListIndex < 0 Then
  c_actividad.ListIndex = 0
End If
End Sub

Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
  c_prov.ListIndex = 0
End If
Call iniciacli
End Sub


Sub inicia()
espere.Show
espere.Label1 = "Inicializando Comprobante....."
espere.Refresh
   t_letra = vta_clientes.t_letrafact
   gcuit = vta_clientes.t_cuit
   c_vend.ListIndex = buscaindice(c_vend, vta_clientes.t_idvend)
   Set cl_compvta = New comprobantes_venta
   cl_compvta.sucursal = Val(c_sucursal)
   cl_compvta.actual (c_tipocomp.ItemData(c_tipocomp.ListIndex))
   cl_compvta.letra = t_letra
   cl_compvta.SACANUMCOMP
   t_numcomp = Format$(cl_compvta.numcomp, "00000000")
   Set cl_compvta = Nothing
   t_cotizacion = para.cotizacion

   If para.calcula_perc_ib = "S" And t_letra = "A" Then
     Set cl_padronib = New padron_ib
     cl_padronib.cuit_texto = gcuit
     cl_padronib.buscar
     t_alicuotaib = Format$(cl_padronib.tasa_percib, "##0.00")
     Select Case cl_padronib.estado_consulta
     Case Is = "OK"
       Label20 = "¡COMPROBANTE SUJETO A PERCEPCION IB! Consulta del Padron de IB Satistactoria"
     Case Is = "NO"
       Label20 = "¡ATENCION! El contribuyente NO se encuentra en el padron, si corresponde debera aplicarle una percpcion de IB del 3%"
     Case Is = "ER"
       Label20 = "¡CUIDADO! Numero de cuit con formato invalido. Padron NO consultado"
     End Select
     Frame11.Visible = True
     
     Set cl_padronib = Nothing
     
     
   Else
     t_alicuotaib = "0.00"
     T_PERCIB = "0.00"
     'gcuit = "0"
   End If
   Call armagrid
   Unload espere

   If Option2 = True Then
      Command8.Enabled = True
   Else
       Command8.Enabled = False
   End If
'Else
'  Unload espere
'  MsgBox ("Error. No se puedo Inicializa el Cliente")
'End If






End Sub

Private Sub c_sucursal_LostFocus()
If c_sucursal.ListIndex < 0 Then
  c_sucursal.ListIndex = buscaindice(c_sucursal, glo.sucursal)
End If
t_sucursal = Format$(c_sucursal, "0000")
t_numcomp = ""
End Sub

Private Sub c_tipocomp_GotFocus()
btnacepta.Enabled = False
End Sub

Private Sub c_tipocomp_LostFocus()
Call iniciacli
End Sub

Private Sub c_vend_LostFocus()
If c_vend.ListIndex < 0 Then
  c_vend.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
vta_ABM_vend.Show
End Sub

Private Sub Command1_LostFocus()
c_vend.clear
Call carga_vendedores(c_vend)
c_vend.ListIndex = 0

End Sub

Private Sub Command2_Click()
vta_ABM_cli.Show
End Sub

Private Sub Command2_LostFocus()
c_prov.clear
Call carga_clientes(c_prov)
c_prov.ListIndex = 0
End Sub

Private Sub Command3_Click()
vta_selremitos.carga
vta_selremitos.Show
End Sub

Private Sub Command4_Click()
Call calculatotales
End Sub

Private Sub Command5_Click()
vta_clientes.t_id = c_prov.ItemData(c_prov.ListIndex)
vta_clientes.carga
vta_clientes.Show
End Sub

Private Sub Command6_Click()
Load gen_consultaib
gen_consultaib.t_id = c_prov.ItemData(c_prov.ListIndex)
gen_consultaib.t_tipo = "C"
gen_consultaib.Show
gen_consultaib.carga
End Sub

Private Sub Command7_Click()
vta_facturacion2.Show
End Sub

Private Sub Command8_Click()
 
  vta_formapago.Show
  vta_formapago.t_total = t_total
 
End Sub

Private Sub Form_Activate()
Frame2.Enabled = False

End Sub

Sub grabaformapago()
  For i = 1 To vta_formapago.msf2.Rows - 1
         If Val(vta_formapago.msf2.TextMatrix(i, 0)) = 3 Then
                'ch. terceros
                q = "select * from cyb_03"
                Set rs = New ADODB.Recordset
                rs.Open q, cn1, adOpenDynamic, adLockOptimistic
                rs.AddNew
                 rs("fecha_emision") = t_fecha
                 rs("num_cheque") = Val(vta_formapago.msf2.TextMatrix(i, 2))
                 rs("banco") = vta_formapago.msf2.TextMatrix(i, 3)
                 rs("sucursal") = vta_formapago.msf2.TextMatrix(i, 4)
                 rs("titular") = vta_formapago.msf2.TextMatrix(i, 5)
                 rs("importe") = Val(vta_formapago.msf2.TextMatrix(i, 6))
                 rs("estado") = "C"
                 rs("fecha_dif") = vta_formapago.msf2.TextMatrix(i, 7)
                 rs("origen") = Left$(vta_clientes.t_cli, 50)
                 rs("destino") = " "
                 rs("num_mov_banco_i") = 0
                 rs("num_mov_banco_e") = 0
                 rs("num_int_op") = 0
                 rs("num_int_rbo") = numint
                 rs("fecha_salida") = t_fecha
                 rs("fecha_ingreso") = t_fecha
                 rs("tipo_salida") = "C"
                rs.Update
                
                qr = "SELECT @@IDENTITY AS NewID"
                Set rs = cn1.Execute(qr)
                numintch = rs.Fields("NewID").Value

                
                Set rs = Nothing
         
         Else
           numintch = 0
         End If
         
         
         If Val(vta_formapago.msf2.TextMatrix(i, 0)) = 4 Then
                q = "select * from cyb_04"
                Set rs = New ADODB.Recordset
                rs.Open q, cn1, adOpenDynamic, adLockOptimistic
                rs.AddNew
                 rs("id_banco") = Val(vta_formapago.msf2.TextMatrix(i, 8))
                 rs("fecha") = vta_formapago.msf2.TextMatrix(i, 7)
                 rs("importe") = Val(vta_formapago.msf2.TextMatrix(i, 6))
                 rs("id_tipomov") = 60 'transf
                 rs("fecha_dif") = vta_formapago.msf2.TextMatrix(i, 7)
                 rs("ubicacion") = "H"
                 rs("entro") = "N"
                 rs("fecha_acreed") = vta_formapago.msf2.TextMatrix(i, 7)
                 rs("num_comp") = Val(vta_formapago.msf2.TextMatrix(i, 2))
                 rs("detalle") = "Transf." & Left$(vta_formapago.msf2.TextMatrix(i, 5), 30)
                 rs("modulo") = "V"
                 rs("num_mov_int") = numint
                 rs("id_tipodbcr") = 1
                rs.Update
                
                Set rs = Nothing
         End If
         
         
         q = "select * from cyb_01 where [id_forma_pago] = " & Val(vta_formapago.msf2.TextMatrix(i, 0))
         Set rs = New ADODB.Recordset
         rs.Open q, cn1
         If Not rs.EOF And Not rs.BOF Then
          If rs("CAJA") = "S" Then
             ctach = rs("id_cuenta_cont")
             QUERY = "INSERT INTO cyb_05([id_cuenta_caja], [id_cuenta_contra], [descripcion], [importe], [ubicacion], [fecha], [num_mov_int], [modulo], [operacion], [id_forma_pago], [num_int_ch_terc])"
             QUERY = QUERY & " VALUES (" & ctach & ", " & para.cuenta_deudores & ", '" & RTrim$(Left$(vta_clientes.t_cli, 49)) & " ', " & Val(vta_formapago.msf2.TextMatrix(i, 6)) & ", 'D', '" & t_fecha & "', " & numint & ", 'V', 'Ctdo " & t_letra & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & "', " & Val(vta_formapago.msf2.TextMatrix(i, 0)) & ", " & numintch & ")"
             cn1.Execute QUERY
          End If
         End If
         Set rs = Nothing

                 
        'formas de pago
        QUERY = "INSERT INTO vta_04([num_int], [secuencia], [id_formapago], [formapago], [num_ch], [detalle_banco], [sucursal], [titular], [importe], [fecha_dif], [num_int_fp])"
        QUERY = QUERY & " VALUES (" & numint & ", " & i & ", " & Val(vta_formapago.msf2.TextMatrix(i, 0)) & ", '" & Left$(RTrim$(vta_formapago.msf2.TextMatrix(i, 1)), 9) & " ', " & Val(vta_formapago.msf2.TextMatrix(i, 2)) & ", '" & RTrim$(vta_formapago.msf2.TextMatrix(i, 3)) & " ', '" & RTrim$(vta_formapago.msf2.TextMatrix(i, 4)) & " ', '" & RTrim$(vta_formapago.msf2.TextMatrix(i, 5)) & " ', " & Val(vta_formapago.msf2.TextMatrix(i, 6)) & ", '" & RTrim$(vta_formapago.msf2.TextMatrix(i, 7)) & " ', " & numintch & ")"
        cn1.Execute QUERY

      Next i

End Sub
Sub iniciacli()
 If c_prov.ListIndex > 0 Then
   vta_clientes.t_id = c_prov.ItemData(c_prov.ListIndex)
   vta_clientes.carga
 Else
   If Val(vta_clientes.t_id) <> 0 Then
      vta_clientes.t_id = 0
      vta_clientes.limpia
   End If
 End If
 
End Sub
Sub actualizaremitos()
For J = 1 To msf1.Rows - 1
 If Val(msf1.TextMatrix(J, 1)) > 1 Then
  cantidadf = Val(msf1.TextMatrix(J, 3)) 'cantidad facturada
  codprodant = Val(msf1.TextMatrix(J, 1))  'cantidad facturada
  i = 1  'para cada articulo busco en los remitos seleccionados
  While i < vta_selremitos.msf1.Rows
   If vta_selremitos.msf1.TextMatrix(i, 0) = "**" Then
     nir = Val(vta_selremitos.msf1.TextMatrix(i, 4))
     q = "SELECT * FROM VTA_02 WHERE [NUM_INT] = " & nir
     Set rs = New ADODB.Recordset
     rs.Open q, cn1, adOpenDynamic, adLockOptimistic
     If Not rs.EOF And Not rs.BOF Then
        'busco el producto en el remito
        q = "select * from vta_03 where [num_int] = " & nir & " and [id_producto] = " & codprodant
        Set rs1 = New ADODB.Recordset
        rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
        While Not rs1.EOF
             'si encontre el producto en el remito
             'verifico cantidad a facturar cantidadf contra lo que hay en el remito
             If cantidadf >= rs1.Fields("Cantidad") Then
                cantidadf = cantidadf - rs1.Fields("Cantidad")
                cpend = 0
                rs1("cantidad") = cpend
                rs1.Update
            
             Else
                cpend = rs1.Fields("cantidad") - cantidadf
                cantidadf = 0
                rs1("cantidad") = cpend
                rs1.Update
                rs1.MoveLast
                i = vta_selremitos.msf1.Rows
             End If
              
            rs1.MoveNext
         Wend
         Set rs1 = Nothing
         
         If verificaremito(nir) = 0 Then
             rs("estado") = "F"
             rs.Update
         End If
        End If
      Set rs = Nothing
    End If
    i = i + 1
  Wend
 End If
Next J
End Sub
Function verificaremito(ByVal n As Long) As Integer
q = "select * from vta_03 where [num_int] = " & n
Set rs1 = New ADODB.Recordset
rs1.Open q, cn1
p = 0
While Not rs1.EOF
  If rs1("id_producto") > 1 Then
    If rs1("cantidad") > 0 Then
      p = 1
    End If
  End If
  rs1.MoveNext
Wend
verificaremito = p
End Function
Sub calculatotales()
vta_facturacion2.armagrid
If t_letra = "A" Then
  s = 0
  V = 0
  For i = 1 To msf1.Rows - 1
      r = Val(msf1.TextMatrix(i, 7))
      s = s + r
      V = V + (r * Val(msf1.TextMatrix(i, 6)) / 100)
      
      'agrega en composicion de iva
      x = 1
      While x < vta_facturacion2.msf1.Rows
        If Val(vta_facturacion2.msf1.TextMatrix(x, 0)) = Val(msf1.TextMatrix(i, 6)) Then
           vta_facturacion2.msf1.TextMatrix(x, 1) = Val(vta_facturacion2.msf1.TextMatrix(x, 1)) + r
           vta_facturacion2.msf1.TextMatrix(x, 2) = Val(vta_facturacion2.msf1.TextMatrix(x, 2)) + (r * Val(msf1.TextMatrix(i, 6)) / 100)
           x = vta_facturacion2.msf1.Rows
        Else
           x = x + 1
        End If
      Wend
  
      
  
  Next i
  vta_facturacion2.sacatotales
  t_subtotal = vta_facturacion2.msf1.TextMatrix(9, 1)
  t_iva = vta_facturacion2.msf1.TextMatrix(9, 2)
  Call sacatotales
  Call sacaperc
  Call sacatotales
 Else
  s = 0
  V = 0
  t = 0
  For i = 1 To msf1.Rows - 1
      r = Val(msf1.TextMatrix(i, 7))
      R2 = Val(msf1.TextMatrix(i, 8))
      s = s + r
      t = t + (R2 * Val(msf1.TextMatrix(i, 3)))
  
            'agrega en composicion de iva
      x = 1
      While x < vta_facturacion2.msf1.Rows
        If Val(vta_facturacion2.msf1.TextMatrix(x, 0)) = Val(msf1.TextMatrix(i, 6)) Then
           vta_facturacion2.msf1.TextMatrix(x, 1) = Val(vta_facturacion2.msf1.TextMatrix(x, 1)) + r
           vta_facturacion2.msf1.TextMatrix(x, 2) = Val(vta_facturacion2.msf1.TextMatrix(x, 2)) + (r * Val(msf1.TextMatrix(i, 6)) / 100)
           x = vta_facturacion2.msf1.Rows
        Else
           x = x + 1
        End If
      Wend
  
  
  Next i
  t_subtotal = s
  t_iva = t - s
  Call sacatotales
  Call sacaperc
  Call sacatotales
 End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     gen_tools.Show
  
End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call TabEnter2(Me, 19)
End If


End Sub

Private Sub Form_Load()
Call INICIALIZA2(Me)
Call carga_clientes(c_prov)
c_prov.AddItem "<Cliente Manual>", 0
c_prov.ListIndex = 0

c_sucursal.AddItem Format$(glo.sucursalf, "0000")
c_sucursal.ListIndex = 0


 Set rs = New ADODB.Recordset
 q = "select * from vta_06 where [sucursal] = " & glo.sucursalf & " and  [id_tipocomp] >= 315 AND [id_tipocomp] <= 330  order by [id_tipocomp]"
 rs.Open q, cn1
 Call llena_combo(rs, "descripcion", "id_tipocomp", c_tipocomp, True)
 Set rs = Nothing
 
Set cl_fiscal = New fiscal
cl_fiscal.carga (glo.sucursalf)
If cl_fiscal.imprimefact = "S" Then
  c_tipocomp.RemoveItem 0
Else
  c_tipocomp.RemoveItem 1
End If
c_tipocomp.ListIndex = 0
Set cl_fiscal = Nothing

Set rs = New ADODB.Recordset
q = "select * from vta_05 order by [denominacion]"
rs.Open q, cn1
Call llena_combo(rs, "denominacion", "id_vendedor", c_vend, True)
Set rs = Nothing
c_vend.ListIndex = 0
Call armagrid
Call barraesag(Me)
Option1 = True
t_sucursal = Format$(glo.sucursalf, "0000")
Load fsc_facturacion1
Load vta_selremitos
Load vta_facturacion2
Frame11.Visible = False

Call carga_actividades(c_actividad)

Load vta_clientes
Load vta_formapago
vta_clientes.limpia
gcuit = "0"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload fsc_facturacion1
Unload vta_facturacion2
Unload vta_selremitos
Unload vta_clientes
Unload vta_formapago
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[INS] Agrega - [ENTER] Modifica - [F5] Elimina - [F9] Graba "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If
Me.KeyPreview = False

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
 If msf1.Rows > 2 Then
    msf1.RemoveItem (msf1.Row)
    
 Else
   Call armagrid
   
 End If
 Call calculatotales
End If

If KeyCode = vbKeyF9 Then
  Call calculatotales
  Call sacatotales
  Call sacaperc
  Call sacatotales

  Frame2.Enabled = True
  btnacepta.Enabled = True
  c_vend.SetFocus
End If

If KeyCode = vbKeyInsert Then
   fsc_facturacion1.t_renglon = ""
   fsc_facturacion1.t_cantidad = ""
   fsc_facturacion1.t_pu = ""
   fsc_facturacion1.t_importe = ""
   fsc_facturacion1.Show
End If

If KeyCode = vbKeyF12 Then
  gen_tools.Show
End If
End Sub

Sub graba()
   'On Error GoTo ERRORGRABA
  
   
  numint = saca_ultnumero_int_comp("V")
      
  Set cl_compvta = New comprobantes_venta
  cl_compvta.sucursal = Val(t_sucursal)
  cl_compvta.actual (c_tipocomp.ItemData(c_tipocomp.ListIndex))
  cl_compvta.letra = t_letra
  cl_compvta.numcomp = Val(t_numcomp)
      
      
      If Option1 = True Then
         ep = "N"
         cp = "0000-00000000"
         contado = "N"
         
      Else
         ep = "S"
         cp = "ctdo"
         contado = "S"
         cl_compvta.ctacte = "N"
      End If
      
      cl_compvta.ACTUALIZA_NUMERADOR
      
      If Option4 = True Then
        moneda = "P"
      Else
        moneda = "D"
      End If
      
      
      Set rs = New ADODB.Recordset
      q = "select * from g8 where [id_actividad] = " & c_actividad.ItemData(c_actividad.ListIndex)
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
       codact = rs("id_actividad")
       alicuotaib = rs("alicuota_ib")
      Else
       codact = 0
       alicuotaib = 0
      End If
      Set rs = Nothing
      
        
      'Set cl_cli = New Clientes
      'cl_cli.carga (c_prov.ItemData(c_prov.ListIndex))
              
      tiporespiva = vta_clientes.c_iva.ItemData(vta_clientes.c_iva.ListIndex)
       
      If c_prov.ListIndex = 0 Then
        idcli = 1
      Else
        idcli = c_prov.ItemData(c_prov.ListIndex)
      End If
      
      cn1.BeginTrans
       
       
       QUERY = "INSERT INTO vta_02([num_int], [sucursal], [num_comp], [letra], [id_tipocomp], [id_cliente], [fecha], [id_usuario], [subtotal], [impuestos], [iva], [total]," & _
"[estado], [id_cuenta], [stock], [cta_cte], [grabado], [estado_pago], [recibo_Pago], [observaciones], [cotizacion_dolar], [total_otra_moneda], [moneda], [id_vendedor], " & _
" [VENTA], [CONTADO], [perc_ib], [perc_gan], [perc_iva] , [id_actividad], [alicuota_ib], [alicuota_perc_iva], [canje_cereal], [fecha_vto], [total_bultos],  [valor_declarado], " & _
" [transporte], [direccion_transp], [cuit_transp], [perc_ss], [sucursal_ingreso], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02])"



QUERY = QUERY & " VALUES (" & numint & ", " & Val(t_sucursal) & ", " & Val(t_numcomp) & ", '" & t_letra & "', " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & _
", " & idcli & ", '" & t_fecha & "', " & para.id_usuario & ", " & Val(t_subtotal) & ", " & Val(t_nograbado) & ", " & Val(t_iva) & ", " & Val(t_total) & _
", 'A', " & para.cuenta_ventas & ", '" & cl_compvta.STOCK & "', '" & cl_compvta.ctacte & "', '" & cl_compvta.grabado & "', '" & ep & "', '" & cp & "', '" & t_observaciones & _
" ', " & Val(t_cotizacion) & ", " & Val(T_total2) & ", '" & moneda & "', " & c_vend.ItemData(c_vend.ListIndex) & ", '" & cl_compvta.venta & "', '" & contado & "', " & Val(t_perc) & _
", 0, " & Val(t_perciva) & ", " & codact & ", " & Val(t_alicuotaib) & ", " & Val(t_alicuotaperciva) & ", " & Check1 & ", '" & t_fechavto & "', 0, 0, ' ', ' ', ' ', 0, " & Val(c_sucursal) & _
", '" & Left$(vta_clientes.t_cli, 50) & "', '" & Left$(vta_clientes.t_direccion, 50) & "', '" & Left$(vta_clientes.t_cuit, 20) & "', '" & Left$(vta_clientes.t_localidad, 50) & "', " & tiporespiva & ")"

                                                                                                                                                                                                                                                            
            
            cn1.Execute QUERY
      
      Set cl_cli = Nothing
      
      For i = 1 To msf1.Rows - 1
        If Val(msf1.TextMatrix(i, 1)) > 1 Then
          Set cl_prod = New productos
          cl_prod.cargar (Val(msf1.TextMatrix(i, 1)))
          costo = cl_prod.precio_ult_compra
          Set cl_prod = Nothing
        Else
          costo = 0
        End If
        
        QUERY = "INSERT INTO vta_03([num_int], [RENGLON], [id_producto], [descripcion], [cantidad], [pu], [importe], [tasaiva], [impuesto], [costo], [cantidad_original], [tunidad], [pu_final])"
        QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i, 0)) & ", " & Val(msf1.TextMatrix(i, 1)) & ", '" & msf1.TextMatrix(i, 2) & " ', " & Val(msf1.TextMatrix(i, 3)) & ", " & Val(msf1.TextMatrix(i, 5)) & ", " & Val(msf1.TextMatrix(i, 7)) & ", " & Val(msf1.TextMatrix(i, 6)) & ", 0, " & costo & ", " & Val(msf1.TextMatrix(i, 3)) & ", '" & msf1.TextMatrix(i, 4) & "', " & Val(msf1.TextMatrix(i, 8)) & ")"
        cn1.Execute QUERY
      
        If cl_compvta.STOCK <> "N" Then
           QUERY = "INSERT INTO stk_01([fecha], [id_producto], [cantidad], [ubicacion], [comprobante], [descripcion], [num_mov_int], [modulo])"
           QUERY = QUERY & " VALUES ('" & t_fecha & "', " & Val(msf1.TextMatrix(i, 1)) & ", " & msf1.TextMatrix(i, 3) & ", '" & cl_compvta.STOCK & "', '" & cl_compvta.abreviatura & t_letra & Format$(t_sucursal, "0000") & "-" & Format$(t_numcomp, "00000000") & "', '" & Left$(c_prov, 50) & "', " & numint & ",'V'" & ")"
           cn1.Execute QUERY
          
           If cl_compvta.STOCK = "E" Then
             c = Val(msf1.TextMatrix(i, 3))
           Else
             c = -Val(msf1.TextMatrix(i, 3))
           End If
           q = "update a2 set [stock] = [stock] + " & c & " where [id_producto] = " & Val(msf1.TextMatrix(i, 1))
           cn1.Execute q
        
        End If
        
        If cl_compvta.venta <> "N" Then
           ultvta = t_letra & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & " | " & Left$(c_prov, 28) & " | " & t_fecha & " | " & Format$(Val(msf1.TextMatrix(i, 4)), "#####0.00")
           QUERY = "update a2 set  [ultima_venta]='" & ultvta & "'"
           QUERY = QUERY & " where [id_producto]= " & Val(msf1.TextMatrix(i, 1))
           cn1.Execute QUERY
        End If
      Next i
      
      
      'actualizo tasa de iva
      For i = 1 To 7
        If Val(vta_facturacion2.msf1.TextMatrix(i, 1)) > 0 Then
          QUERY = "INSERT INTO vta_09([num_int], [tasa_iva], [iva], [neto], [tipo_iva])"
          QUERY = QUERY & " VALUES (" & numint & ", " & Val(vta_facturacion2.msf1.TextMatrix(i, 0)) & ", " & Val(vta_facturacion2.msf1.TextMatrix(i, 2)) & ", " & Val(vta_facturacion2.msf1.TextMatrix(i, 1)) & ", " & tiporespiva & ")"
          cn1.Execute QUERY
          
        End If
      Next i
     
     
     If Option2 = True Then
        'graba fortma de pago
        Call grabaformapago
        
      End If
      

     If cl_compvta.contabilidad <> "N" Then
         numintcgr = saca_ultnumero_int_comp("G")

         If Option1 = True Then
           cta = para.cuenta_deudores
         Else
           cta = para.cuenta_caja
         End If
         u1 = cl_compvta.contabilidad
          
         If u1 = "D" Then
           u2 = "H"
         Else
           u2 = "D"
         End If
         
         Set rs = New ADODB.Recordset
         q = "select * from c_01 where [id_cuenta] = " & cta
         rs.Open q, cn1
         If Not rs.EOF And Not rs.BOF Then
           dcta = rs("descripcion")
         Else
           dcta = "Cuenta Inexistente"
         End If
         Set rs = Nothing
         
         'grabo asiento
         QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Ventas] " & cl_compvta.abreviatura & " " & t_letra & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & "', 'V', " & numint & ", " & Val(t_total) & ", " & Val(t_total) & ", " & para.id_usuario & ", '" & Left$(RTrim$(c_prov), 50) & "')"
         cn1.Execute QUERY
      
         ic = 1
         'cuenta madre ctacte o caja
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cta & ", '" & u1 & "', " & Val(t_total) & ", '" & dcta & "')"
         
         cn1.Execute QUERY
         ic = ic + 1
      
         If Val(t_nograbado) > 0 Then
           'cuenta nogbra
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_conceptos_nograbados & ", '" & u2 & "', " & Val(t_nograbado) & ", 'No Grabado')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
                   
         If Val(t_perc) > 0 Then
           'cuenta perc
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_perc_IB & ", '" & u2 & "', " & Val(t_perc) & ", 'Perc. IB')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
          
          If Val(t_perciva) > 0 Then
           'cuenta perc
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_perc_iva & ", '" & u2 & "', " & Val(t_perciva) & ", 'Perc. IVA')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
         If Val(t_iva) > 0 Then
           'cuenta perc
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_iva_ventas & ", '" & u2 & "', " & Val(t_iva) & ", 'IVA')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
         
         'contrapartida
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_ventas & ", '" & u2 & "', " & Val(t_subtotal) & ", '" & "Ventas" & "')"
         cn1.Execute QUERY
      
      End If
      
      cn1.CommitTrans
      Set rs = Nothing
      Set cl_compvta = Nothing
      Set cl_cli = Nothing

      
      
      
      'actualizo remitos
     'If Val(vta_selremitos.t_seleccionados) > 0 Then
        For i = 1 To vta_selremitos.msf1.Rows - 1
          If vta_selremitos.msf1.TextMatrix(i, 0) = "**" Then
             QUERY = "INSERT INTO vta_08([id_factura], [id_remito])"
             QUERY = QUERY & " VALUES (" & numint & ", " & Val(vta_selremitos.msf1.TextMatrix(i, 4)) & ")"
             cn1.Execute QUERY
          End If
        Next i
     
        Call actualizaremitos
     'End If
      
      J = MsgBox("Confirma Impresion del Comprobante", 4)
      If J = 6 Then
           Set cl_compvta = New comprobantes_venta
           cl_compvta.cargar2 (numint)
           cl_compvta.imprimir
      End If
      Call INICIALIZA2(Me)
      Call armagrid
      c_tipocomp.SetFocus
      Frame2.Enabled = False
      t_sucursal = Format$(c_sucursal, "0000")
      vta_formapago.armagrid2
      
Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos y vuelva a repetir la operacion")
  

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    fsc_facturacion1.t_renglon = msf1.Row
    fsc_facturacion1.t_basico = msf1.TextMatrix(msf1.Row, 1)
    fsc_facturacion1.t_detalle = msf1.TextMatrix(msf1.Row, 2)
    fsc_facturacion1.t_cantidad = msf1.TextMatrix(msf1.Row, 3)
    fsc_facturacion1.t_unidad = msf1.TextMatrix(msf1.Row, 4)
    fsc_facturacion1.t_pu = msf1.TextMatrix(msf1.Row, 5)
    fsc_facturacion1.t_importe = msf1.TextMatrix(msf1.Row, 7)
    fsc_facturacion1.Show
  End If
End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight
Me.KeyPreview = True
Call calculatotales
End Sub

Private Sub Option1_Click()
Command8.Enabled = False
End Sub

Private Sub Option1_GotFocus()
Call keyform(Me, "D")

End Sub

Private Sub Option1_LostFocus()
Call keyform(Me, "A")

End Sub

Private Sub Option2_Click()
Command8.Enabled = True
End Sub

Private Sub Option2_GotFocus()
'all keyform(Me, "A")

End Sub

Private Sub Option2_LostFocus()
'Call keyform(Me, "D")

End Sub


Private Sub t_alicuotaib_LostFocus()
If Val(t_alicuotaib) < 0 Then
  t_alicuotaib = "0.00"
End If
Call sacaperc
End Sub

Private Sub t_cotizacion_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)
End Sub

Private Sub t_cotizacion_LostFocus()
If Val(t_cotizacion) <= 0 Then
   t_cotizacion = 1
End If
End Sub

Private Sub t_fecha_LostFocus()
If Not IsDate(t_fecha) Then
  t_fecha = Format$(Now, "dd/mm/yyyy")
Else
  t_fecha = Format$(t_fecha, "dd/mm/yyyy")
End If
End Sub


Private Sub t_fechavto_LostFocus()
If Not IsDate(t_fechavto) Then
  t_fechavto = Format$(Now, "dd/mm/yyyy")
Else
  t_fechavto = Format$(t_fechavto, "dd/mm/yyyy")
End If

End Sub

Private Sub t_iva_LostFocus()
Call sacatotales

End Sub

Private Sub t_nograbado_LostFocus()
Call sacatotales

End Sub


Private Sub t_numcomp_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)

End Sub

Private Sub t_numcomp_LostFocus()
If IsNumeric(t_numcomp) Then
   t_numcomp = Format$(t_numcomp, "00000000")
   Call carga
Else
  t_numcomp.SetFocus
End If
End Sub

Private Sub t_observaciones_LostFocus()
Call NULOS(t_observaciones)
End Sub

Private Sub t_perc_LostFocus()
Call sacatotales

End Sub

Private Sub t_subtotal_LostFocus()
Call sacatotales
End Sub
Sub sacatotales()
t_subtotal = Format$(Val(t_subtotal), "######0.00")
t_nograbado = Format$(Val(t_nograbado), "######0.00")
t_perc = Format$(Val(t_perc), "######0.00")
t_iva = Format$(Val(t_iva), "######0.00")
t_perciva = Format$(Val(t_perciva), "######0.00")
t_total = Format$(Val(t_subtotal) + Val(t_nograbado) + Val(t_perc) + Val(t_iva) + Val(t_perciva), "######0.00")
If Option4 = True Then
 If Val(t_cotizacion) < 1 Then
   t_cotizacion = 1
 End If
 T_total2 = Format$(Val(t_total) / Val(t_cotizacion), "#####0.00")
Else
  T_total2 = Format$(Val(t_total) * Val(t_cotizacion), "#####0.00")
End If
End Sub
Sub sacaperc()
If Option3 = True Then
   s = Val(t_subtotal) * Val(t_cotizacion)
 Else
   s = Val(t_subtotal)
 End If

q = "select * from i_01 where [id_impuesto] = 1"
Set rs2 = New ADODB.Recordset
rs2.Open q, cn1
If Not rs2.EOF And Not rs2.BOF Then
 impmin = rs2("importe_minimo_sujeto_ret")
 retmin = rs2("retencion-minima")
  If rs2("calcula") = "S" And s >= impmin Then
   tp = s * (Val(t_alicuotaib) / 100)
   If tp >= retmin Then
       t_perc = Format$(tp, "#####0.00")
   Else
       t_perc = "0.00"
   End If
 Else
  t_perc = "0.00"
 End If
Else
 t_perc = "0.00"
End If

If Option3 = True Then 'dolares
   p$ = Val(t_perc) / Val(t_cotizacion)
Else
   p$ = Val(t_perc)
End If
t_perc = Format$(p$, "####0.00")
Set rs2 = Nothing

If Check1 = 1 Then
  'calcula perciva rg 2459
   Set rs2 = New ADODB.Recordset
   q = "select * from  i_01, i_02 where i_01.[id_impuesto] = i_02.[id_impuesto] and i_01.[id_impuesto] = 2"
   rs2.Open q, cn1
   If Not rs2.EOF And Not rs2.BOF Then
     'Set cl_cli = New Clientes
     cl_cli.carga (c_prov.ItemData(c_prov.ListIndex))
     If cl_cli.id > 0 Then
        If cl_cli.operador_granos = "S" Then
           t_alicuotaperciva = rs2("alicuota_i")
        Else
           t_alicuotaperciva = rs2("alicuota_n")
        End If
     Else
        t_alicuotaperciva = "0.00"
     End If
    ' Set cl_cli = Nothing
   Else
     t_alicuotaperciva = "0.00"
   End If
   
   
   If s >= rs2("importe_minimo_sujeto_ret") Then
     pi = s * Val(t_alicuotaperciva) / 100
   Else
     pi = 0
   End If
   
   If Option3 = True Then 'dolares
    p$ = pi / Val(t_cotizacion)
   Else
    p$ = pi
   End If
           
   If p$ >= rs2("retencion-minima") Then
      t_perciva = Format$(pi, "######0.00")
   Else
      t_perciva = "0.00"
   End If
   Set rs2 = Nothing
Else
  t_alicuotaperciva = "0.00"
  t_perciva = "0.00"
End If




End Sub

Private Sub t_sucursal_GotFocus()
t_sucursal = Format$(Val(c_sucursal), "0000")
End Sub

Private Sub t_sucursal_LostFocus()
Call inicia
End Sub

Private Sub t_total_LostFocus()
t_total = Format$(t_total, "######0.00")
End Sub

Private Sub T_total2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Option2 = True Then
      Command8.Enabled = True
   Else
      Command8.Enabled = False
 End If
 btnacepta.Enabled = True
 btnacepta.SetFocus
End If

End Sub
