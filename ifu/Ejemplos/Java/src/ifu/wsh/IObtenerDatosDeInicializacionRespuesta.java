package ifu.wsh  ;

import com4j.*;

@IID("{44C8E088-C222-4FC1-94ED-9395F5FE32C2}")
public interface IObtenerDatosDeInicializacionRespuesta extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "NroCUIT"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(201) //= 0xc9. The runtime will prefer the VTID if present
  @VTID(7)
  java.lang.String nroCUIT();


  /**
   * <p>
   * Getter method for the COM property "RazonSocial"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(202) //= 0xca. The runtime will prefer the VTID if present
  @VTID(8)
  java.lang.String razonSocial();


  /**
   * <p>
   * Getter method for the COM property "NroSerie"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(203) //= 0xcb. The runtime will prefer the VTID if present
  @VTID(9)
  java.lang.String nroSerie();


  /**
   * <p>
   * Getter method for the COM property "FechaInicializacion"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(204) //= 0xcc. The runtime will prefer the VTID if present
  @VTID(10)
  java.lang.String fechaInicializacion();


  /**
   * <p>
   * Getter method for the COM property "NroPOS"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(205) //= 0xcd. The runtime will prefer the VTID if present
  @VTID(11)
  java.lang.String nroPOS();


  /**
   * <p>
   * Getter method for the COM property "FechaIniActividades"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(206) //= 0xce. The runtime will prefer the VTID if present
  @VTID(12)
  java.lang.String fechaIniActividades();


  /**
   * <p>
   * Getter method for the COM property "CodIngBrutos"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(207) //= 0xcf. The runtime will prefer the VTID if present
  @VTID(13)
  java.lang.String codIngBrutos();


  /**
   * <p>
   * Getter method for the COM property "RespIVA"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(208) //= 0xd0. The runtime will prefer the VTID if present
  @VTID(14)
  java.lang.String respIVA();


  /**
   * <p>
   * Getter method for the COM property "Resultado"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(209) //= 0xd1. The runtime will prefer the VTID if present
  @VTID(15)
  boolean resultado();


  // Properties:
}
