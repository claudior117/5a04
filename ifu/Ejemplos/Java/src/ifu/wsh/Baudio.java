package ifu.wsh  ;

import com4j.*;

/**
 */
public enum Baudio implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2400
   * </p>
   */
  bd2400(2400),
  /**
   * <p>
   * The value of this constant is 4800
   * </p>
   */
  bd4800(4800),
  /**
   * <p>
   * The value of this constant is 9600
   * </p>
   */
  bd9600(9600),
  /**
   * <p>
   * The value of this constant is 19200
   * </p>
   */
  bd19200(19200),
  /**
   * <p>
   * The value of this constant is 38400
   * </p>
   */
  bd38400(38400),
  /**
   * <p>
   * The value of this constant is 57600
   * </p>
   */
  bd57600(57600),
  /**
   * <p>
   * The value of this constant is 115200
   * </p>
   */
  bd115200(115200),
  ;

  private final int value;
  Baudio(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
