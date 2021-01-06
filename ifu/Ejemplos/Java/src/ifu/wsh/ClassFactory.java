package ifu.wsh  ;

import com4j.*;

/**
 * Defines methods to create COM objects
 */
public abstract class ClassFactory {
  private ClassFactory() {} // instanciation is not allowed


  /**
   * Driver Object
   */
  public static ifu.wsh.IDriver createDriver() {
    return COM4J.createInstance(ifu.wsh.IDriver.class, "{536413FB-C017-4B59-8923-AE79800E3BB4}" );
  }

  public static ifu.wsh.IObtenerDatosDeInicializacionRespuesta createObtenerDatosDeInicializacionRespuesta() {
    return COM4J.createInstance(ifu.wsh.IObtenerDatosDeInicializacionRespuesta.class, "{EF88ACD1-CD97-418F-A01B-B4657E28C6B2}" );
  }

  public static ifu.wsh.ISubtotalRespuesta createSubtotalRespuesta() {
    return COM4J.createInstance(ifu.wsh.ISubtotalRespuesta.class, "{27D2653D-A3D2-4037-A5AD-EF73A64A0C69}" );
  }

  public static ifu.wsh.ICierreZTotales createCierreZTotales() {
    return COM4J.createInstance(ifu.wsh.ICierreZTotales.class, "{F0C532B6-9FDC-4A80-BEC1-C9A064F5400D}" );
  }
}
