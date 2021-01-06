package ifu.wsh  ;

import com4j.*;

/**
 */
public enum ModeloPrn implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  modHasar715(0),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  modHasar715v2(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  modHasar615(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  modHasar320(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  modHasarPR4F(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  modHasarPR5F(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  modHasar950(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  modHasar951(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  modHasar441(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  modHasar321(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  modHasar322(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  modHasar322v2(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  modHasar330(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  modHasar1120(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  modHasarPL8F(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  modHasarPL8Fv2(16),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  modHasarPL23(17),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  modEpsonTM300AF(18),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  modEpsonTMU220AF(19),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  modEpsonTM2000(20),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  modEpsonTM2000AFPlus(21),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  modEpsonLX300(22),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  modHasarPT1000F(23),
  /**
   * <p>
   * The value of this constant is 24
   * </p>
   */
  modEpsonTMT900FA(24),
  ;

  private final int value;
  ModeloPrn(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
