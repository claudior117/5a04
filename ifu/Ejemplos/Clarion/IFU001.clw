

   MEMBER('IFU.clw')                                       ! This is a MEMBER module


   INCLUDE('ABTOOLBA.INC'),ONCE
   INCLUDE('ABWINDOW.INC'),ONCE

                     MAP
                       INCLUDE('IFU001.INC'),ONCE        !Local module procedure declarations
                     END


Main PROCEDURE                                             ! Generated from procedure template - Frame

ResultOk             ULONG                                 !
Precio               REAL                                  !
Cantidad             REAL                                  !
Monto                REAL                                  !
Texto                STRING(20)                            !
OLE                  LONG                                  !
DisplayDayString STRING('Sunday   Monday   Tuesday  WednesdayThursday Friday   Saturday ')
DisplayDayText   STRING(9),DIM(7),OVER(DisplayDayString)
AppFrame             APPLICATION('Process Customer Orders'),AT(,,400,253),FONT('MS Sans Serif',8,COLOR:Black,),STATUS(-1,80,120,45),SYSTEM,MAX,RESIZE,IMM
                       TOOLBAR,AT(0,0,400,253)
                         BUTTON('Cierre Z'),AT(19,10,76,25),USE(?Button2)
                         BUTTON('Factura B'),AT(17,50,77,27),USE(?Button1)
                       END
                     END

ThisWindow           CLASS(WindowManager)
Ask                    PROCEDURE(),DERIVED                 ! Method added to host embed code
Init                   PROCEDURE(),BYTE,PROC,DERIVED       ! Method added to host embed code
Kill                   PROCEDURE(),BYTE,PROC,DERIVED       ! Method added to host embed code
TakeAccepted           PROCEDURE(),BYTE,PROC,DERIVED       ! Method added to host embed code
TakeWindowEvent        PROCEDURE(),BYTE,PROC,DERIVED       ! Method added to host embed code
                     END

Toolbar              ToolbarClass

  CODE
  GlobalResponse = ThisWindow.Run()                        ! Opens the window and starts an Accept Loop

!---------------------------------------------------------------------------
DefineListboxStyle ROUTINE
!|
!| This routine create all the styles to be shared in this window
!| It's called after the window open
!|
!---------------------------------------------------------------------------

ThisWindow.Ask PROCEDURE

  CODE
  IF NOT INRANGE(AppFrame{Prop:Timer},1,100)
    AppFrame{Prop:Timer} = 100
  END
    AppFrame{Prop:StatusText,3} = CLIP(DisplayDayText[(TODAY()%7)+1]) & ', ' & FORMAT(TODAY(),@D4)
    AppFrame{Prop:StatusText,4} = FORMAT(CLOCK(),@T3)
  PARENT.Ask


ThisWindow.Init PROCEDURE

ReturnValue          BYTE,AUTO

  CODE
  GlobalErrors.SetProcedureName('Main')
  SELF.Request = GlobalRequest                             ! Store the incoming request
  ReturnValue = PARENT.Init()
  IF ReturnValue THEN RETURN ReturnValue.
  SELF.FirstField = 1
  SELF.VCRRequest &= VCRRequest
  SELF.Errors &= GlobalErrors                              ! Set this windows ErrorManager to the global ErrorManager
  CLEAR(GlobalRequest)                                     ! Clear GlobalRequest after storing locally
  CLEAR(GlobalResponse)
  SELF.AddItem(Toolbar)
  OPEN(AppFrame)                                           ! Open window
  SELF.Opened=True
  Do DefineListboxStyle
  INIMgr.Fetch('Main',AppFrame)                            ! Restore window settings from non-volatile store
  SELF.SetAlerts()
  RETURN ReturnValue


ThisWindow.Kill PROCEDURE

ReturnValue          BYTE,AUTO

  CODE
  ReturnValue = PARENT.Kill()
  IF ReturnValue THEN RETURN ReturnValue.
  IF SELF.Opened
    INIMgr.Update('Main',AppFrame)                         ! Save window data to non-volatile store
  END
  GlobalErrors.SetProcedureName
  RETURN ReturnValue


ThisWindow.TakeAccepted PROCEDURE

ReturnValue          BYTE,AUTO

Looped BYTE
  CODE
  LOOP                                                     ! This method receive all EVENT:Accepted's
    IF Looped
      RETURN Level:Notify
    ELSE
      Looped = 1
    END
  ReturnValue = PARENT.TakeAccepted()
    CASE ACCEPTED()
    OF ?Button2
          OLE = Create(0, Create:OLE)
          OLE{PROP:Create} = 'IFUniversal.Driver'
          OLE{'Depurar'} = 1
          OLE{'Modelo'} = 23
          OLE{'Puerto'} = 31
          OLE{'Baudios'} = 9600
      
          ResultOk = OLE{'Inicializar()'}
      
          IF ResultOK <> 0
            OLE{'CancelarComprobante()'}
          END
      
          IF ResultOK <> 0
            ResultOk = OLE{'CierreZ()'}
          END
      
      
        IF ResultOk <> 0
          MESSAGE('Cierre realizado exitosamente!')
        ELSE
          MESSAGE(OLE{'ErrorDesc'})
        END
    OF ?Button1
      ! Ejemplo de impresion de una factura B usando el driver IF Universal
      ! Documentacion en linea sobre metodos y parametros en http://bitingenieria.com.ar/doc/ifu/IFUniversal_TLB/IDriver.html
      ! Ver Constantes.txt para entender los valores de los parametros
      
          OLE = Create(0, Create:OLE)
          OLE{PROP:Create} = 'IFUniversal.Driver'
          OLE{'Depurar'} = 1
          OLE{'Modelo'} = 23     ! Modelo HasarPT1000 2G
          OLE{'Puerto'} = 31
          OLE{'Baudios'} = 9600
      
          OLE{'Inicializar()'}
          OLE{'CancelarComprobante()'}
      
          ResultOk = 1
      
          ! Esto no se envia si la factura es a consumidor final
          If not ResultOk Then
      !*        function DatosCliente(const aNombre: WideString; aTipoDeDocumento: TipoDeDocumento;
      !*                              const aDocumento: WideString; aResponsIVA: ResponsabilidadIVA;
      !*                              const aDireccion: WideString): OLE_CANCELBOOL;
             ResultOk = OLE{'DatosCliente("Abel Miranda", 0, "20939802593", 1, "Blanco Encalada 1204 5to A")'}
          End
          IF ResultOk <> 0
            ResultOk = OLE{'AbrirComprobante(2)'}
          END
             
          IF ResultOk <> 0
             ResultOk = OLE{'ImprimirTextoFiscal(COD. ARICULO:000001)'}
          End
      
          IF ResultOk <> 0
      !*        function ImprimirItem2g(Descripcion: WideString; Cantidad: Double; Precio: Double; IVA: Double;
      !*                                ImpuestosInternos: Double; g2CondicionIVA: CondicionesIVA ; g2TipoImpuestoInterno: WideString;
      !*                                g2UnidadReferencia: Integer; g2CodigoProducto: WideString; g2CodigoInterno: WideString;
      !*                                g2UnidadMedida: UnidadesMedida )
             BIND('Cantidad', Cantidad)
             BIND('Precio', Precio)
             Cantidad =1
             Precio = 0.1
             ResultOk = OLE{'ImprimirItem2g("Item 1", Cantidad, Precio, 21, 0, 7, 0, 1, "7790001001054", "", 7)'}
          END
      
          IF ResultOk <> 0
      !*        function ImprimirDescuentoGeneral(const Descripcion: WideString; Monto: Double): OLE_CANCELBOOL;
             BIND('Monto', Monto)
             Monto = 0.05
             ResultOk = OLE{'ImprimirDescuentoGeneral("Descuento", Monto)'}
          END
      
          IF ResultOk <> 0
             ResultOk = OLE{'ImprimirPago2g("Efectivo", 1, "", 8, 1, "", "")'}
          END
      
          IF ResultOk <> 0
            OLE{'CerrarComprobante()'}
          END
      
        IF ResultOk <> 0
          MESSAGE('Comprobante impreso exitosamente!')
        ELSE
          MESSAGE(OLE{'ErrorDesc'})
        END
    END
    RETURN ReturnValue
  END
  ReturnValue = Level:Fatal
  RETURN ReturnValue


ThisWindow.TakeWindowEvent PROCEDURE

ReturnValue          BYTE,AUTO

Looped BYTE
  CODE
  LOOP                                                     ! This method receives all window specific events
    IF Looped
      RETURN Level:Notify
    ELSE
      Looped = 1
    END
  ReturnValue = PARENT.TakeWindowEvent()
    CASE EVENT()
    OF Event:Timer
      AppFrame{Prop:StatusText,3} = CLIP(DisplayDayText[(TODAY()%7)+1]) & ', ' & FORMAT(TODAY(),@D4)
      AppFrame{Prop:StatusText,4} = FORMAT(CLOCK(),@T3)
    ELSE
    END
    RETURN ReturnValue
  END
  ReturnValue = Level:Fatal
  RETURN ReturnValue

