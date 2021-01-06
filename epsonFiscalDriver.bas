Attribute VB_Name = "epsonFiscalDriver"
'/******************************************************************************
'*   Autor: Business Support And Development Unit                              *
'*                                                                             *
'*   Este código es gratuito y permite integrar impresoras fiscales EPSON      *
'*   usando la librería de bajo nivel (EpsonFiscalDriver.dll).                 *
'*                                                                             *
'*   Para implementarla las funciones de la libreria (.dll) usted debe incluir *
'*   el modulo 'EpsonFiscalDriver' en su proyecto.                             *
'*                                                                             *
'*   Este software se entrega con fines didácticos y sin garantía alguna.      *
'*   EPSON  NO ASUME responsabilidad legal alguna.                             *
'*   El programador usa este software bajo su propio riesgo y responsabilidad. *
'******************************************************************************/


 Public Declare Sub setProtocolType Lib "EpsonFiscalDriver.dll" (ByVal Protocol As Long)
 Public Declare Function getProtocolType Lib "EpsonFiscalDriver.dll" () As Long
 Public Declare Sub setComPort Lib "EpsonFiscalDriver.dll" (ByVal portnumber As Long)
 Public Declare Function getComPort Lib "EpsonFiscalDriver.dll" () As Long
 Public Declare Sub setBaudRate Lib "EpsonFiscalDriver.dll" (ByVal baud As Long)
 Public Declare Function getBaudRate Lib "EpsonFiscalDriver.dll" () As Long
 Public Declare Sub OpenPort Lib "EpsonFiscalDriver.dll" ()
 Public Declare Sub OpenPortByName Lib "EpsonFiscalDriver.dll" (ByVal Port As String)

 Public Declare Sub ClosePort Lib "EpsonFiscalDriver.dll" ()
 Public Declare Sub Purge Lib "EpsonFiscalDriver.dll" ()

 Public Declare Function SetLog Lib "EpsonFiscalDriver.dll" (ByVal filePath As String, ByVal bUserAction As Boolean)
 Public Declare Function getState Lib "EpsonFiscalDriver.dll" () As Long
 Public Declare Function getLastError Lib "EpsonFiscalDriver.dll" () As Long

 Public Declare Function getFiscalStatus Lib "EpsonFiscalDriver.dll" () As Long
 Public Declare Function getPrinterStatus Lib "EpsonFiscalDriver.dll" () As Long
 Public Declare Function getReturnCode Lib "EpsonFiscalDriver.dll" () As Long

 Public Declare Function getExtraFieldCount Lib "EpsonFiscalDriver.dll" () As Long
 
 
 Public Declare Sub GetAPIVersion Lib "EpsonFiscalDriver.dll" (ByVal output_buffer As String)
 
 
 Public Declare Sub AddDataField Lib "EpsonFiscalDriver.dll" (ByVal buffer As String, ByVal buffer_length As Long)
 Public Declare Sub SendCommand Lib "EpsonFiscalDriver.dll" ()
 Public Declare Sub GetExtraField Lib "EpsonFiscalDriver.dll" (ByVal field_number As Long, ByVal output_buffer As String, ByVal output_buffer_length As Long, ByVal output_buffer_final_length As Long)
    
  
 Public Declare Sub GetSentFrame Lib "EpsonFiscalDriver.dll" (ByRef output_buffer As String, ByVal output_buffer_length As Long, ByRef output_buffer_final_length As Long)
 Public Declare Sub GetReceivedFrame Lib "EpsonFiscalDriver.dll" (ByRef output_buffer As String, ByVal output_buffer_length As Long, ByRef output_buffer_final_length As Long)
 Public Declare Sub SendCommandComplete Lib "EpsonFiscalDriver.dll" (ByVal command As String)


