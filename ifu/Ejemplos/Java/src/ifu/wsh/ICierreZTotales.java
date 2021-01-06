package ifu.wsh  ;

import com4j.*;

@IID("{A7973DAB-A411-454D-927E-517037721A21}")
public interface ICierreZTotales extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "FNDTotalVentas"
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(201) //= 0xc9. The runtime will prefer the VTID if present
  @VTID(7)
  double fndTotalVentas();


  /**
   * <p>
   * Getter method for the COM property "FNDTotalIVA"
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(202) //= 0xca. The runtime will prefer the VTID if present
  @VTID(8)
  double fndTotalIVA();


  /**
   * <p>
   * Getter method for the COM property "FNDTotalImpuestosInternos"
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(203) //= 0xcb. The runtime will prefer the VTID if present
  @VTID(9)
  double fndTotalImpuestosInternos();


  /**
   * <p>
   * Getter method for the COM property "FNDTotalOtrosTributos"
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(204) //= 0xcc. The runtime will prefer the VTID if present
  @VTID(10)
  double fndTotalOtrosTributos();


  /**
   * <p>
   * Getter method for the COM property "NCTotalVentas"
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(205) //= 0xcd. The runtime will prefer the VTID if present
  @VTID(11)
  double ncTotalVentas();


  /**
   * <p>
   * Getter method for the COM property "NCTotalIVA"
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(206) //= 0xce. The runtime will prefer the VTID if present
  @VTID(12)
  double ncTotalIVA();


  /**
   * <p>
   * Getter method for the COM property "NCTotalImpuestosInternos"
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(207) //= 0xcf. The runtime will prefer the VTID if present
  @VTID(13)
  double ncTotalImpuestosInternos();


  /**
   * <p>
   * Getter method for the COM property "NCTotalOtrosTributos"
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(208) //= 0xd0. The runtime will prefer the VTID if present
  @VTID(14)
  double ncTotalOtrosTributos();


  /**
   * <p>
   * Getter method for the COM property "NroCierre"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(209) //= 0xd1. The runtime will prefer the VTID if present
  @VTID(15)
  int nroCierre();


  // Properties:
}
