VERSION 5.00
Object = "{0A6BE9FC-5039-11D5-98EC-0800460222F0}#1.0#0"; "IFEpson.ocx"
Begin VB.Form fsc_errorfiscal 
   Caption         =   "ESTADO IMPRESORA FISCAL"
   ClientHeight    =   8145
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "Tipo Informe"
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton Command1 
         Caption         =   "Verifica"
         Height          =   375
         Left            =   3960
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox c_tipo 
         Height          =   315
         ItemData        =   "fsc002.frx":0000
         Left            =   120
         List            =   "fsc002.frx":0002
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Fecha y Hora en Impresora Fiscal"
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   7200
      Width           =   3255
      Begin VB.TextBox t_hora 
         Height          =   405
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         Height          =   405
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      Top             =   7320
      Width           =   735
   End
   Begin EPSON_Impresora_Fiscal.PrinterFiscal epson2 
      Left            =   360
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame Frame4 
      Caption         =   "Descripcion"
      Height          =   3615
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   4815
      Begin VB.ListBox List1 
         Height          =   3180
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Estado Impresor Fiscal"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   2760
      TabIndex        =   4
      Top             =   2280
      Width           =   2175
      Begin VB.TextBox t_ib 
         Height          =   405
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox t_ih 
         Alignment       =   2  'Center
         Height          =   405
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Estado Placa Fiscal"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
      Begin VB.TextBox t_fb 
         Height          =   405
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox t_fh 
         Alignment       =   2  'Center
         Height          =   405
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ESTADO"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   2775
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   120
         Picture         =   "fsc002.frx":0004
         Top             =   240
         Width           =   795
      End
   End
End
Attribute VB_Name = "fsc_errorfiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub c_tipo_LostFocus()
If c_tipo.ListIndex < 0 Then
   c_tipo.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
Set cl_fiscal = New fiscal
cl_fiscal.carga (glo.sucursalf)
If cl_fiscal.idmodelo <> 24 Then 'tm-900 then

        epson2.PortNumber = cl_fiscal.puerto
        Select Case c_tipo.ListIndex
         Case Is = 0
            'normal
            Call verifica
          Case Is = 1
            'controlador
            Call controlador
          Case Is = 2
           'contribuyente
            Call contribuyente
          Case Is = 3
           'contribuyente
            Call contadores
            
        
        End Select

Else
  MsgBox ("Este controlador fiscal no tiene implementada esta funcion")
End If
Set cl_fiscal = Nothing

 



End Sub
Sub contribuyente()
List1.clear
t_fh = "0000"
t_ih = "0000"
r = epson2.Status("C")
t_fh = epson2.FiscalStatus
t_ih = epson2.PrinterStatus

If Not r Then
  Text1 = "Error"
  Image1.Picture = LoadPicture(App.Path & "\tools\error.jpg")
  List1.AddItem "--------------------------------------------------------"
  List1.AddItem "　　ERROR!!!!!   Impresora Apagada o sin Conexion"
  List1.AddItem "--------------------------------------------------------"
 
  
Else
    Image1.Picture = LoadPicture(App.Path & "\tools\ok.jpg")
    Text1 = "OK"
    Call verificaerrfiscal(epson2.FiscalStatus, epson2.PrinterStatus)
    nc = epson2.AnswerField_3
    PVt = epson2.AnswerField_4
    ti = epson2.AnswerField_5
    mtf = epson2.AnswerField_7
    
    List1.AddItem " "
    List1.AddItem " "
    List1.AddItem "Datos Obtenidos del Impresor"
    List1.AddItem " "
    List1.AddItem "Cuit Contribuyente................>  " & nc
    List1.AddItem "Punto de Venta....................>  " & PVt
    List1.AddItem "Tipo de Iva.......................>  " & ti
    List1.AddItem "Monto maximo tique factura........>  " & mtf

    
    r = epson2.SetGetDateTime("G")
    If r Then
       t_fecha = Mid$(epson2.AnswerField_3, 5, 2) & "/" & Mid$(epson2.AnswerField_3, 3, 2) & "/" & Mid$(epson2.AnswerField_3, 1, 2)
       t_hora = Mid$(epson2.AnswerField_4, 1, 2) & ":" & Mid$(epson2.AnswerField_4, 3, 2)
    Else
       t_fecha = ""
       t_hora = ""
    End If
  End If

End Sub

Sub contadores()
List1.clear
t_fh = "0000"
t_ih = "0000"
r = epson2.Status("A")
t_fh = epson2.FiscalStatus
t_ih = epson2.PrinterStatus
't_fb = HEXABIN(t_fh)
't_ib = HEXABIN(t_ih)

If Not r Then
  
  Text1 = "Error"
  Image1.Picture = LoadPicture(App.Path & "\tools\error.jpg")
  List1.AddItem "--------------------------------------------------------"
  List1.AddItem "　　ERROR!!!!!   Impresora Apagada o sin Conexion"
  List1.AddItem "--------------------------------------------------------"
 
  
Else
    Image1.Picture = LoadPicture(App.Path & "\tools\ok.jpg")
    Text1 = "OK"
    Call verificaerrfiscal(epson2.FiscalStatus, epson2.PrinterStatus)
    
    uz = epson2.AnswerField_3
    utb = epson2.AnswerField_4
    uti = epson2.AnswerField_5
    uab = epson2.AnswerField_6
    uai = epson2.AnswerField_7
    
    List1.AddItem " "
    List1.AddItem " "
    List1.AddItem "Datos Obtenidos del Impresor"
    List1.AddItem " "
    List1.AddItem "Ultimo Cierre Z................................>  " & uz
    List1.AddItem "Ult. tique o fact. B emitido sin problemas.....>  " & utb
    List1.AddItem "Ult. tique o fact. b impreso...................>  " & uti
    List1.AddItem "Ult.  fact. A emitido sin problemas............>  " & uab
    List1.AddItem "Ult.  fact. A impreso..........................>  " & uai
        
    r = epson2.SetGetDateTime("G")
    If r Then
       t_fecha = Mid$(epson2.AnswerField_3, 5, 2) & "/" & Mid$(epson2.AnswerField_3, 3, 2) & "/" & Mid$(epson2.AnswerField_3, 1, 2)
       t_hora = Mid$(epson2.AnswerField_4, 1, 2) & ":" & Mid$(epson2.AnswerField_4, 3, 2)
    Else
       t_fecha = ""
       t_hora = ""
    End If
  End If


End Sub
Sub verifica()
List1.clear
t_fh = "0000"
t_ih = "0000"
r = epson2.Status("N")

t_fh = epson2.FiscalStatus
t_ih = epson2.PrinterStatus
't_fb = HEXABIN(t_fh)
't_ib = HEXABIN(t_ih)

If Not r Then
  Text1 = "Error"
  Image1.Picture = LoadPicture(App.Path & "\tools\error.jpg")
  List1.AddItem "--------------------------------------------------------"
  List1.AddItem "　　ERROR!!!!!   Impresora Apagada o sin Conexion"
  List1.AddItem "--------------------------------------------------------"
  
Else
    Call verificaerrfiscal(epson2.FiscalStatus, epson2.PrinterStatus)
    
    Image1.Picture = LoadPicture(App.Path & "\tools\ok.jpg")
    Text1 = "OK"
    
    nc = epson2.AnswerField_3
    f1c = epson2.AnswerField_4
    h1c = epson2.AnswerField_5
    uz = epson2.AnswerField_6
    ap = epson2.AnswerField_7
    at = epson2.AnswerField_4

    List1.AddItem " "
    List1.AddItem " "
    List1.AddItem "Datos Obtenidos del Impresor"
    List1.AddItem " "
    List1.AddItem "Ult. Comprobante Fiscal Emitido............>  " & nc
    List1.AddItem "Fecha primer comp. jornada fiscal..........>  " & f1c
    List1.AddItem "Hora primer comp. jornada fiscal...........>  " & h1c
    List1.AddItem "Nro. Ultimo cierre Z.......................>  " & uz
    List1.AddItem "Dato Auditoria Parcial.....................>  " & ap
    List1.AddItem "Dato Auditoria Total.......................>  " & at

    
    r = epson2.SetGetDateTime("G")
    If r Then
       t_fecha = Mid$(epson2.AnswerField_3, 5, 2) & "/" & Mid$(epson2.AnswerField_3, 3, 2) & "/" & Mid$(epson2.AnswerField_3, 1, 2)
       t_hora = Mid$(epson2.AnswerField_4, 1, 2) & ":" & Mid$(epson2.AnswerField_4, 3, 2)
    Else
       t_fecha = ""
       t_hora = ""
    End If
  End If

End Sub

Sub controlador()
List1.clear
t_fh = "0000"
t_ih = "0000"
r = epson2.Status("P")
t_fh = epson2.FiscalStatus
t_ih = epson2.PrinterStatus
  
If Not r Then
  
  Text1 = "Error"
  Image1.Picture = LoadPicture(App.Path & "\tools\error.jpg")
  List1.AddItem "--------------------------------------------------------"
  List1.AddItem "　　ERROR!!!!!   Impresora Apagada o sin Conexion"
  List1.AddItem "--------------------------------------------------------"
 
  
Else
    Image1.Picture = LoadPicture(App.Path & "\tools\ok.jpg")
    Text1 = "OK"
    Call verificaerrfiscal(epson2.FiscalStatus, epson2.PrinterStatus)
    
    ac = epson2.AnswerField_6
    IT = epson2.AnswerField_8
    itf = epson2.AnswerField_9
    iff = epson2.AnswerField_10
    m = epson2.AnswerField_13

    List1.AddItem " "
    List1.AddItem " "
    List1.AddItem "Datos Obtenidos del Impresor"
    List1.AddItem " "
    List1.AddItem "Ancho en Columnas.....................>  " & ac
    List1.AddItem "Imprime Tique.........................>  " & IT
    List1.AddItem "Imprime Tique Factura.................>  " & itf
    List1.AddItem "Imprime Factura.......................>  " & iff
    List1.AddItem "Modelo Impresora Fiscal...............>  " & m

    
    r = epson2.SetGetDateTime("G")
    If r Then
       t_fecha = Mid$(epson2.AnswerField_3, 5, 2) & "/" & Mid$(epson2.AnswerField_3, 3, 2) & "/" & Mid$(epson2.AnswerField_3, 1, 2)
       t_hora = Mid$(epson2.AnswerField_4, 1, 2) & ":" & Mid$(epson2.AnswerField_4, 3, 2)
    Else
       t_fecha = ""
       t_hora = ""
    End If
  End If

End Sub
Function buscaerrorf() As Integer
  'devuelve 1 si no hay error y 0 si lo hay
  
  Dim ef(16) As Integer 'error placa fiscal
  b = 16
  Errorf = 1 '1 = ok    0 = error
  m = 0
  msg1 = "　LLAME AL PROGRAMADOR!!"
  msg2 = "　LLAME AL TECNICO FISCAL!!"
  For i = 0 To 15
    ef(i) = Val(Mid$(t_fb, b - i, 1)) 'pone el estado de bit 0, 1, n
  Next i
    
  'VERIFICACION PLACA FISCAL
  'verificacion rapida del bit 15 de la placa fiscal
  List1.AddItem "----------------------------------------------------------------------------"
  If ef(15) = 1 Then
    Errorf = 0
    List1.AddItem "ESTADO PLACA FISCAL: 　　 ERROR !!!!"
  Else
    List1.AddItem "ESTADO PLACA FISCAL: ### O.K. ###"
  End If
  List1.AddItem "----------------------------------------------------------------------------"
  List1.AddItem " "
  If ef(0) = 1 Then 'BIT 0
    Errorf = 0
    List1.AddItem "[Bit 0] --> Error Comprobancion Memoria Fiscal "
    List1.AddItem "            " & msg2
  End If
  If ef(1) = 1 Then
    Errorf = 0
    List1.AddItem "[Bit 1] --> Error Comprobancion Memoria de Trabajo "
    List1.AddItem "            " & msg2
  End If
  If ef(2) = 1 Then
    Errorf = 0
    List1.AddItem "[Bit 2] --> Poca Bateria "
    List1.AddItem "            En corto tiempo se bloqueara la impresora"
  End If
  If ef(3) = 1 Then
    Errorf = 0
    List1.AddItem "[Bit 3] --> Comando no reconocido "
    List1.AddItem "            " & msg1
  End If
  If ef(4) = 1 Then
    Errorf = 0
    List1.AddItem "[Bit 4] --> Campo de Datos Invalido "
    List1.AddItem "            " & msg1
  End If
  If ef(5) = 1 Then
    Errorf = 0
    List1.AddItem "[Bit 5] --> Comando no Valido para Estado Fiscal "
    List1.AddItem "            " & msg1
  End If

  If ef(6) = 1 Then
    Errorf = 0
    List1.AddItem "[Bit 6] --> Desbordamiento de Totales "
    List1.AddItem "            " & msg1
  End If

  If ef(7) = 1 Then
    Errorf = 0
    List1.AddItem "[Bit 7] --> Memoria Fiscal Llena "
    List1.AddItem "            " & msg2
  End If

  If ef(8) = 1 Then
    Errorf = 0
    List1.AddItem "[Bit 8] --> Memoria Fiscal Casi LLena "
    List1.AddItem "            " & msg2
  End If

  If ef(11) = 1 Then
    Errorf = 0
    List1.AddItem "[Bit11] --> Es Nenesario Realizar un Cierre Z "
  End If

   buscaerrorf = Errorf
  
End Function

Function buscaerrori() As Integer
  'devuelve 1 si no hay error y 0 si lo hay
  
  Dim ei(16) As Integer 'error impresora fiscal
  b = 16
  Errori = 1 '1 = ok    0 = error
  m = 0
  msg1 = "　LLAME AL PROGRAMADOR!!"
  msg2 = "　LLAME AL TECNICO FISCAL!!"
  For i = 0 To 15
    ei(i) = Val(Mid$(t_ib, b - i, 1)) 'pone el estado de bit 0, 1, n
  Next i
    
  'VERIFICACION impresora
  'verificacion rapida del bit 15 de la placa fiscal
  List1.AddItem "----------------------------------------------------------------------------"
  If ei(15) = 1 Then
    Errori = 0
    List1.AddItem "ESTADO IMPRESORA: 　　 ERROR !!!!"
  Else
    List1.AddItem "ESTADO IMPRESORA: ### O.K. ###"
  End If
  List1.AddItem "----------------------------------------------------------------------------"
  List1.AddItem " "
  If ei(0) = 1 Then 'BIT 0
    Errori = 0
    List1.AddItem "[Bit 0] --> No se usa "
  End If
  If ei(1) = 1 Then
    Errori = 0
    List1.AddItem "[Bit 1] --> No se Usa "
  End If
  If ei(2) = 1 Then
    Errori = 0
    List1.AddItem "[Bit 2] --> Error y / o Falla de la Impresora"
    List1.AddItem "            " & msg2
  End If
  If ei(3) = 1 Then
    Errori = 0
    List1.AddItem "[Bit 3] --> Impresora Fuera de Linea "
  End If
  If ei(4) = 1 Then
    Errori = 0
    List1.AddItem "[Bit 4] --> Poco Papel cinta Auditoria "
  End If
  If ei(5) = 1 Then
    Errori = 0
    List1.AddItem "[Bit 5] --> Poco Papel para Comprobantes "
  End If

  If ei(6) = 1 Then
    Errori = 0
    List1.AddItem "[Bit 6] --> Buffer de Impresora LLeno "
  End If

  If ei(14) = 1 Then
    Errori = 0
    List1.AddItem "[Bit14] --> Sin Papel"
  End If


  buscaerrori = Errori
  
End Function

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
c_tipo.AddItem "N - > Normal", 0
c_tipo.AddItem "P - > Caracteristicas del Controlador", 1
c_tipo.AddItem "C - > Informacion del Titular ", 2
c_tipo.AddItem "A - > Contadores de doc. fiscales", 3
c_tipo.ListIndex = 0





End Sub
