VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "epsonFiscalTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

' -----------------------------------------------------------------------------
' Descripción: Son utitlidades generales que permiten ver la implementación
'              de las funciones de la libreria  "EpsonFiscalDriver.dll"
' -----------------------------------------------------------------------------

  ' -------------------------------------------------------
  '   Constantes usadas en la clase
  '--------------------------------------------------------
  Public ERROR_NONE As Long '0
  Public ERROR_PORT_USED As Long '= '3
  Public ERROR_PORT_CLOSED As Long '= 6
  Public ERROR_INVALID_DATA_FIELD As Long '= 8
  Public ERROR_OPERATION_PROGRESS As Long '= 9
  Public ERROR_TIMEOUT As Long '= 11
  Public ERROR_INVALID_INDEX As Long '= 13
  Public ERROR_INVALID_FRAME_COMMAND As Long '= 14
  Public ERROR_INVALID_DOC As Long '= 15
  
  
  ' status
  Public EFPROT_CLOSED As Long '= 0
  Public EFPROT_IDLE As Long '= 1
  Public EFPROT_BUSY As Long '= 2
 
  'constantes
  Public MAX_DESCRIPTION As Long '= 200
  Public SIZE_API_VERSION As Long '= 10

  
  ' protocolos soportados
  Public ProtocoloCompatible As Long '= 0
  Public ProtocoloExtendido As Long '= 1


  ' Listado de errores
  Private errorList(15) As String
  'errorList = ["Exito","Puerto en uso","Numero de puerto invalido","Puerto abierto","Puerto cerrado", \
  '              "Error de inicializacion","Puerto cerrado","Maximo tamaño alcanzado","Campo de datos invalido","Opracion en progreso", \
  '              "Error de escritura", "Error en tiempo de espera","respuesta de rechazo recibida (NACK '0x15')", "Index invalido", \
  '              "datos de comando invalido"]

  
  ' -------------------------------------------------------
  '               Variables usadas en la clase
  '--------------------------------------------------------
  Private portnumber As String '= "0"  ' USB
  Private baudRate As String '= 9600
  Private protocolo As String  '= ProtocoloExtendido
  Private habIP As Boolean


 ' *******************************************************************************************************************
  '                                                Métodos privados
  ' *******************************************************************************************************************
  ' -----------------------------------------------------------------------------
  ' Carga los datos separados por | para ser enviados a la impresora
  ' No se puede cargar mas de 1 byte a la vez debido a que si algun byte es
  ' valor 0 se genera error ya que es interpretado como final de cadena
  ' -----------------------------------------------------------------------------
  Public Function tools_load_command(ByVal cmd As String) As Long
    Dim cad As String
    Dim fin As Long
    Dim i As Long
    Dim tam As Long
    
    Call Purge
    
    cad = ""
    fin = Len(cmd)
    i = 0
    Do While (i < fin)
    
      If (Mid(cmd, i, 1) = "|") Then
        tam = Len(cad)
        Call AddDataField(cad, tam)
        cad = ""
        i = i + 1
        If (i >= fin) Then
          Exit Do
        End If
      End If

      If (Mid(cmd, i, 1) = "|") Then
          cad = ""
          Call AddDataField(cad, 0)
          i = i + 1
          If (i >= fin) Then
            Exit Do
          End If
      Else
        cad = cad + Mid(cmd, i, 1)
      End If
      
      i = i + 1
    Loop
    
   
    If (cad = "") Then
      Call AddDataField(cad, 0)
    Else
      tam = Len(cad)
      Call AddDataField(cad, tam)
    End If
    
    tools_load_command = tools_send_command()
      
  End Function
  
      
 ' -----------------------------------------------------------------------------
  ' Envio de datos almacenados en memoria de dll  con retorno de error
  ' -----------------------------------------------------------------------------
  Public Function tools_send_command() As Long
    Dim error As Long
    Dim time As Long
    Dim repeatTime As Byte
    Dim Count As Byte
    Dim start As Long
    Dim currentTime As Long
    
    
    
    Do While (True)

      ' envio de comando
      Call SendCommand
      

      ' validacion interna en la dll previa al envio de comando
      error = getLastError()
      If Not (error = ERROR_NONE) Then
        Exit Do
      End If
    

      'Asegura tiempo de espera de retorno del retorno del comando

        '-------- VARIABLES DE CONFIGURACION -----
        time = 3 ' segundos
        repeatTime = 3 ' cantidad de reintentos
        '-----------------------------------------

        Count = 0
        start = Timer
        Do While (Count < repeatTime)
          
            If (getState() = EFPROT_IDLE) Then
                Exit Do
            End If
        
            currentTime = Timer - start
            If (currentTime >= time) Then
                start = Timer
                Count = Count + 1
            End If
        Loop


      'define error por timeout
      If (Count = repeatTime) Then
        error = ERROR_TIMEOUT
        Exit Do
      End If


      ' Verifica error asociado al comando enviado
      error = getLastError()
      If (error = ERROR_NONE) Then
        error = getReturnCode()
      End If
           
      Exit Do
    Loop
      
    tools_send_command = error
  End Function


  ' -----------------------------------------------------------------------------
  ' Envio de comandos de cadena completa "sendCommandComplete" con retorno de error
  ' -----------------------------------------------------------------------------
  Public Function tools_send_command_string(ByVal cmd As String) As Long
    Dim error As Long

    Do While (True)
      If (cmd = "") Then
        error = ERROR_INVALID_FRAME_COMMAND
        Exit Do
      End If
      

      ' limpia buffer
      Call Purge
      ' envio de cadena
      Call SendCommandComplete(cmd)
      'obtención del error
      error = getLastError()
      If (error = ERROR_NONE) Then
        error = getReturnCode()
      End If
      
      Exit Do
    Loop

    tools_send_command_string = error
  End Function
   
   
   
  ' -----------------------------------------------------------------------------
  ' Permite enviar el comando y obtener la respuesta del error
  ' Esta funcion es debido a que no es soportado el x'HH'
  ' -----------------------------------------------------------------------------
  Public Function tools_send_data(ByVal cmd As String) As Long
    Dim error As Long
      
    If (protocolo = ProtocoloCompatible) Then
      error = tools_load_command(cmd)
    Else
      error = tools_send_command_string(cmd)
    End If
    
    
    tools_send_data = error
  End Function
   
   
  ' *******************************************************************************************************************
  '                                                Métodos publicos
  ' *******************************************************************************************************************


  ' -----------------------------------------------------------------------------
  ' Incializacion de las variables del set de comandos
  ' -----------------------------------------------------------------------------
    Private Sub Class_Initialize()
        ' -------------------------------------------------------
        '               Constantes usadas en la clase
        '--------------------------------------------------------
         ERROR_NONE = 0
         ERROR_PORT_USED = 3
         ERROR_PORT_CLOSED = 6
         ERROR_INVALID_DATA_FIELD = 8
         ERROR_OPERATION_PROGRESS = 9
         ERROR_TIMEOUT = 11
         ERROR_INVALID_INDEX = 13
         ERROR_INVALID_FRAME_COMMAND = 14
         ERROR_INVALID_DOC = 15
  
  
         'Status
         EFPROT_CLOSED = 0
         EFPROT_IDLE = 1
         EFPROT_BUSY = 2
 
         MAX_DESCRIPTION = 200
         SIZE_API_VERSION = 10

    
         'protocolos soportados
         ProtocoloCompatible = 0
         ProtocoloExtendido = 1


        ' Listado de errores
        errorList(0) = "Exito"
        errorList(1) = "Puerto en uso"
        errorList(2) = "Numero de puerto invalido"
        errorList(3) = "Puerto abierto"
        errorList(4) = "Puerto cerrado"
        errorList(5) = "Error de inicializacion"
        errorList(6) = "Puerto cerrado"
        errorList(7) = "Maximo tamaño alcanzado"
        errorList(8) = "Campo de datos invalido"
        errorList(9) = "Operacion en progreso"
        errorList(10) = "Error de escritura"
        errorList(11) = "Error en tiempo de espera"
        errorList(12) = "respuesta de rechazo recibida (NACK '0x15')"
        errorList(13) = "Index invalido"
        errorList(14) = "datos de comando invalido"
        
    
  
        ' -------------------------------------------------------
        '               Variables usadas en la clase
        '--------------------------------------------------------
        portnumber = "0"  ' USB
        baudRate = 9600
        protocolo = ProtocoloExtendido
        
    End Sub
    
' -----------------------------------------------------------------------------
' Muestra la descripcion del error
' -----------------------------------------------------------------------------
Public Function ConsultarDescripcionDeError() As String
    Dim cad As String
    Dim retorno As Long

    ' retorno asociado al ENVIO del comando
    retorno = getLastError()
    cad = errorList(retorno)

    If (retorno = ERROR_NONE) Then
        retorno = getReturnCode()
        If Not (retorno = ERROR_NONE) Then
            cad = "Codigo de retorno(" & Hex(retorno) & "). Ver manual de la impresora (Codigos de retorno)."
        End If
    End If

    ConsultarDescripcionDeError = cad
End Function
  
 ' -----------------------------------------------------------------------------
  ' Inica archivo Log
  ' -----------------------------------------------------------------------------
  Public Function ComenzarLog()
    Call SetLog("./", 1)
  End Function
  
  ' -----------------------------------------------------------------------------
  ' Detener archivo Log
  ' -----------------------------------------------------------------------------
  Public Function DetenerLog()
    Call SetLog("./", 0)
  End Function
  
  ' -----------------------------------------------------------------------------
  ' Configura Velocidad por defecto 9600
  ' -----------------------------------------------------------------------------
  Public Function ConfigurarVelocidad(velocidad As Long) As Long
    baudRate = velocidad
    Call setBaudRate(baudRate)

    ConfigurarVelocidad = ERROR_NONE
  End Function
  
  ' -----------------------------------------------------------------------------
  ' Obetner Velocidad ajustada en la clase para la comunicación
  ' -----------------------------------------------------------------------------
  Public Function ObtenerVelocidad() As Long
    ObtenerVelocidad = baudRate
  End Function
  
  ' -----------------------------------------------------------------------------
  ' Obetner Velocidad configurada en la impresora fiscal
  ' -----------------------------------------------------------------------------
  Public Function ObtenerVelocidadFP() As Long
    ObtenerVelocidadFP = getBaudRate()
  End Function
  


  ' -----------------------------------------------------------------------------
  ' Configura Puerto de comunicación por defecto "0" - USB
  ' -----------------------------------------------------------------------------
  Public Function ConfigurarPuerto(ByVal puerto As String) As Long
    
    portnumber = Trim(puerto)

    If InStr(1, puerto, ".") > 0 Then
        habIP = True
    Else
        habIP = False
        Call setComPort(Int(portnumber))
    End If
  

    ConfigurarPuerto = ERROR_NONE
  End Function
  
  ' -----------------------------------------------------------------------------
  ' Configura Protocolo por defecto Extendido
  ' -----------------------------------------------------------------------------
  Public Function ConfigurarProtocolo(ByVal tipoProtocolo As String) As Long
    ' if (tipoProtocolo != self.ProtocoloCompatible) and (tipoProtocolo != self.ProtocoloExtendido):  ' validación
    '   return  self.ERROR_INVALID_DATA_FIELD

    protocolo = tipoProtocolo
    Call setProtocolType(protocolo)

    ConfigurarProtocolo = ERROR_NONE
  End Function
  
  ' -----------------------------------------------------------------------------
  ' Obetner Velocidad ajustada en la clase para la comunicación
  ' -----------------------------------------------------------------------------
  Public Function ObtenerProtocolo() As Long
    ObtenerProtocolo = protocolo
  End Function
  
  ' -----------------------------------------------------------------------------
  ' Obetner Velocidad configurada en la impresora fiscal
  ' -----------------------------------------------------------------------------
  Public Function ObtenerProtocoloFP() As Long
    ObtenerProtocoloFP = getProtocolType()
  End Function
  
   
   ' -----------------------------------------------------------------------------
  ' Permite conectarse al equipo y definir el modelo para ajustar los comandos
  ' y extensiones que serán usados posteriormente
  ' -----------------------------------------------------------------------------
  Public Function Conectar() As Long
    Dim error As Long
    
  
    'Call setProtocolType(protocolo)
    'Call setBaudRate(baudRate)
    'Call setComPort(portnumber)

    Call Purge
    Call OpenPort
    error = getLastError()
    Call Purge

    Conectar = error
  End Function
  
  
  
  
     ' -----------------------------------------------------------------------------
  ' Permite conectarse al equipo y definir el modelo para ajustar los comandos
  ' y extensiones que serán usados posteriormente
  ' -----------------------------------------------------------------------------
  Public Function NewConectar() As Long
    Dim error As Long
    Dim Port As String
  
    'Call setProtocolType(protocolo)
    'Call setBaudRate(baudRate)
   
   
    If (habIP) Then
        Port = "lan:" & portnumber
    Else
        If (portnumber = "0") Then   ' USB
            Port = "usb:USB"
        Else         ' SERIAL
            Port = "serial:COM" & portnumber
        End If
    End If
   
    Call Purge
    Call OpenPortByName(Port)
    error = getLastError()
    Call Purge

    NewConectar = error
  End Function
  
  ' -----------------------------------------------------------------------------
  ' Cierra el puerto de comunicación
  ' -----------------------------------------------------------------------------
  Public Function Desconectar() As Long
    Dim error As Long
    
    Call ClosePort
    error = getLastError()
    Desconectar = error
  End Function
 

  ' -----------------------------------------------------------------------------
  ' Permite enviar el comando y obtener la respuesta del error
  ' en los fields de cada comando no admite valor \x00
  ' -----------------------------------------------------------------------------
  Public Function EnviarComando(ByVal cmd As String) As Long
    Dim error As Long
    
  
    '---------------------------------------------
    ' Se implementa esta funcion ya que no soporta x'HH'
    ' y el protocolo compatible un solo byte para comando y otro para extension
    '---------------------------------------------
    error = tools_send_data(cmd)
     
    '---------------------------------------------
    ' En caso de soportar x'HH' se implementa es funcion
    ' sin diferenciar el tipo de protocolo
    '---------------------------------------------
    ' error = tools_send_command_string(cmd)
    '---------------------------------------------
    EnviarComando = error

  End Function

 
' -----------------------------------------------------------------------------
' Obtiene los datos de retorno en funcion del campo especificado
' -----------------------------------------------------------------------------
Public Function ObtenerRespuestaExtendida(ByVal numero_campo As Long) As String
  
    Dim buffer_salida As String
    Dim largo_buffer_salida As Long
    Dim largo_final_buffer_salida As Long
      
    largo_buffer_salida = 51200 'tamaño de buffer en manual de especificaciones de EpsonFiscalDriver
    buffer_salida = String(largo_buffer_salida, " ")
        
    Call GetExtraField(numero_campo, buffer_salida, largo_buffer_salida, VarPtr(largo_final_buffer_salida))
    buffer_salida = Trim(buffer_salida)
              
    ObtenerRespuestaExtendida = buffer_salida
    
End Function
  
' -----------------------------------------------------------------------------
' Obtiene la verion de la dll
' -----------------------------------------------------------------------------
Public Function ConsultarVersionDll() As String
    Dim buffer_salida As String
    Dim cad As String
    
    ' inicializar variable
    buffer_salida = String(SIZE_API_VERSION, " ")
  
    Call GetAPIVersion(buffer_salida)
    cad = Trim(buffer_salida)
    
    ConsultarVersionDll = cad
   
End Function
  
 
' -----------------------------------------------------------------------------
' Obtiene la verion de la impresora fiscal
' -----------------------------------------------------------------------------
Public Function ConsultarVersionIF(ByVal cmd As String, ByVal num_camp_ver As Long, ByVal num_cam_ver_may As Long, ByVal num_cam_ver_men As Long)
    Dim retorno As Long
    Dim cad As String


    cad = ""
    Do While (True)
        ' solicita información
        retorno = tools_send_data(cmd)

        If Not (retorno = ERROR_NONE) Then
            Exit Do
        End If

        'version
        If Not (num_camp_ver = 0) Then
            cad = "Versión IF: " & ObtenerRespuestaExtendida(num_camp_ver)
        End If


            
        ' version mayor
        If Not (num_cam_ver_may = 0) Then
            cad = cad & ", Versión Mayor: " & ObtenerRespuestaExtendida(num_cam_ver_may)
        End If


        ' version menor
        If Not (num_cam_ver_men = 0) Then
            cad = cad & ", Versión Menor: " & ObtenerRespuestaExtendida(num_cam_ver_men)
        End If


        Exit Do
    Loop

    ConsultarVersionIF = cad
End Function
 
 
  ' -----------------------------------------------------------------------------
  ' Obtiene la verion de la dll
  ' -----------------------------------------------------------------------------
  Public Function ConsultarEstadoImpresora() As Long
    ConsultarEstadoImpresora = getPrinterStatus()
   End Function
 
 
   
' -----------------------------------------------------------------------------
' Obtiene el estado fiscal
' -----------------------------------------------------------------------------
Public Function ConsultarEstadoFiscal() As Long
    ConsultarEstadoFiscal = getFiscalStatus()
End Function
   
   
   
' -----------------------------------------------------------------------------
' Consultar ultimo error
' -----------------------------------------------------------------------------
Public Function ultimoError() As String
    Dim retorno As Long

    retorno = getLastError()
    If (retorno = ERROR_NONE) Then
        retorno = getReturnCode()
        ultimoError = Hex(retorno)
    Else
        ultimoError = Str(retorno)
    End If

 
End Function
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   

