package org.extract.text.analyser.word;

import java.math.BigDecimal;

/**
 * @ClassName SpaceHandler
 * @Description 虽然现在没有用到，但是以后可能会用到的工具类，word中各种距离的转换
 *              比如图片的大小会是emu单位，可以用这个工具类转换为pt
 * @Author WANGHAN756
 * @Date 2021/7/29 10:26
 * @Version 1.0
 **/
public class SpaceHandler {
    public static final double PT_PER_PX = 0.75;
    public static final int IN_PER_PT = 72;
    public static final double CM_PER_PT = 28.3;
    public static final double MM_PER_PT = 2.83;
    public static final int EMU_PER_PX = 9525;


    public static int pxToEMU(double px) {
        return DoubleUtils.mul(px, EMU_PER_PX).intValue();
    }

    public static int emuToPx(double emu) {
        return DoubleUtils.div(emu, EMU_PER_PX).intValue();
    }

    public static double ptToPx(double pt) {
        return DoubleUtils.div(pt, PT_PER_PX);
    }

    public static double inToPx(double in) {
        return DoubleUtils.mul(inToPt(in), PT_PER_PX);
    }

    public static double pxToIn(double px) {
        return ptToIn(pxToPt(px));
    }

    public static double cmToPx(double cm) {
        return DoubleUtils.mul(cmToPt(cm), PT_PER_PX);
    }

    public static double pxToCm(double px) {
        return ptToCm(pxToPt(px));
    }

    public static double mmToPx(double mm) {
        return DoubleUtils.mul(mmToPt(mm), PT_PER_PX);
    }

    public static double pxToMm(double px) {
        return ptToMm(pxToPt(px));
    }

    public static double ptToIn(double pt) {
        return DoubleUtils.div(pt, IN_PER_PT);
    }

    public static double ptToMm(double mm) {
        return DoubleUtils.div(mm, MM_PER_PT);
    }

    public static double ptToCm(double in) {
        return DoubleUtils.div(in, CM_PER_PT);
    }

    public static double pxToPt(double px) {
        return DoubleUtils.mul(px, PT_PER_PX);
    }

    public static double inToPt(double in) {
        return DoubleUtils.mul(in, IN_PER_PT);
    }

    public static double mmToPt(double mm) {
        return DoubleUtils.mul(mm, MM_PER_PT);
    }

    public static double cmToPt(double cm) {
        return DoubleUtils.mul(cm, CM_PER_PT);
    }
}

class DoubleUtils {
    // 默认除法运算精度
    private static final Integer DEF_DIV_SCALE = 2;

    /**
     * 提供精确的加法运算。
     *
     * @param value1
     *            被加数
     * @param value2
     *            加数
     * @return 两个参数的和
     */
    public static Double add(Number value1, Number value2) {
        return add(value1, value2, DEF_DIV_SCALE);
    }

    /**
     * 提供精确的加法运算。
     *
     * @param value1
     *            被加数
     * @param value2
     *            加数
     * @param scale
     *            表示需要精确到小数点以后几位。
     * @return 两个参数的和
     */
    public static Double add(Number value1, Number value2, Integer scale) {
        BigDecimal b1 = new BigDecimal(Double.toString(value1.doubleValue()));
        BigDecimal b2 = new BigDecimal(Double.toString(value2.doubleValue()));
        BigDecimal add = b1.add(b2);
        BigDecimal setScale = add.setScale(scale, BigDecimal.ROUND_HALF_UP);
        return setScale.doubleValue();
    }

    /**
     * 提供精确的减法运算。
     *
     * @param value1
     *            被减数
     * @param value2
     *            减数
     * @return
     */
    public static Double sub(Number value1, Number value2) {
        return sub(value1, value2, DEF_DIV_SCALE);
    }

    /**
     *
     * @param value1
     *            被减数
     * @param value2
     *            减数
     * @param scale
     *            表示需要精确到小数点以后几位。
     * @return 两个参数的差
     */
    public static Double sub(Number value1, Number value2, Integer scale) {
        BigDecimal b1 = new BigDecimal(Double.toString(value1.doubleValue()));
        BigDecimal b2 = new BigDecimal(Double.toString(value2.doubleValue()));
        BigDecimal subtract = b1.subtract(b2);
        BigDecimal setScale = subtract.setScale(scale, BigDecimal.ROUND_HALF_UP);
        return setScale.doubleValue();
    }

    /**
     * 提供精确的乘法运算。
     *
     * @param value1
     *            被乘数
     * @param value2
     *            乘数
     * @return 两个参数的积
     */
    public static Double mul(Number value1, Number value2) {
        return mul(value1, value2, DEF_DIV_SCALE);
    }

    /**
     * 提供精确的乘法运算。
     *
     * @param value1
     *            被乘数
     * @param value2
     *            乘数
     * @param scale
     *            表示需要精确到小数点以后几位。
     * @return 两个参数的积
     */
    public static Double mul(Number value1, Number value2, Integer scale) {
        BigDecimal b1 = new BigDecimal(Double.toString(value1.doubleValue()));
        BigDecimal b2 = new BigDecimal(Double.toString(value2.doubleValue()));
        BigDecimal multiply = b1.multiply(b2);
        BigDecimal setScale = multiply.setScale(scale, BigDecimal.ROUND_HALF_UP);
        return setScale.doubleValue();
    }

    /**
     * 提供精确的乘法运算。
     *
     * @param value1
     *            被乘数
     * @param value2
     *            乘数
     * @param scale
     *            表示需要精确到小数点以后几位。
     *
     * @return 两个参数的积
     */
    public static Double mul(Number value1, Number value2, int scale, int roundingMode) {
        BigDecimal b1 = new BigDecimal(Double.toString(value1.doubleValue()));
        BigDecimal b2 = new BigDecimal(Double.toString(value2.doubleValue()));
        BigDecimal multiply = b1.multiply(b2);
        BigDecimal setScale = multiply.setScale(scale, roundingMode);
        return setScale.doubleValue();
    }

    /**
     * 提供（相对）精确的除法运算，当发生除不尽的情况时， 精确到小数点以后2位，以后的数字四舍五入。
     *
     * @param dividend
     *            被除数
     * @param divisor
     *            除数
     * @return 两个参数的商
     */
    public static Double div(Number dividend, Number divisor) {
        return DoubleUtils.div(dividend, divisor, DEF_DIV_SCALE);
    }

    /**
     * 提供（相对）精确的除法运算。 当发生除不尽的情况时，由scale参数指定精度，以后的数字四舍五入。
     *
     * @param dividend
     *            被除数
     * @param divisor
     *            除数
     * @param scale
     *            表示需要精确到小数点以后几位。
     * @return 两个参数的商
     */
    public static Double div(Number dividend, Number divisor, Integer scale) {
        if (scale < 0) {
            throw new IllegalArgumentException("The scale must be a positive integer or zero");
        }
        BigDecimal b1 = new BigDecimal(Double.toString(dividend.doubleValue()));
        BigDecimal b2 = new BigDecimal(Double.toString(divisor.doubleValue()));
        return b1.divide(b2, scale, BigDecimal.ROUND_HALF_UP).doubleValue();
    }

    /**
     * 提供（相对）精确的除法运算。 当发生除不尽的情况时，由scale参数指定精度，以后的数字四舍五入。
     *
     * @param dividend
     *            被除数
     * @param divisor
     *            除数
     * @param scale
     *            表示需要精确到小数点以后几位。
     * @return 两个参数的商
     */
    public static Double div(Number dividend, Number divisor, Integer scale, int roundingMode) {
        if (scale < 0) {
            throw new IllegalArgumentException("The scale must be a positive integer or zero");
        }
        BigDecimal b1 = new BigDecimal(Double.toString(dividend.doubleValue()));
        BigDecimal b2 = new BigDecimal(Double.toString(divisor.doubleValue()));
        return b1.divide(b2, scale, roundingMode).doubleValue();
    }

    /**
     * 提供精确的小数位四舍五入处理。
     *
     * @param value
     *            需要四舍五入的数字
     * @param scale
     *            小数点后保留几位
     * @return 四舍五入后的结果
     */
    public static Double round(Double value, Integer scale) {
        if (scale < 0) {
            throw new IllegalArgumentException("The scale must be a positive integer or zero");
        }
        BigDecimal b = new BigDecimal(Double.toString(value));
        BigDecimal one = new BigDecimal("1");
        return b.divide(one, scale, BigDecimal.ROUND_HALF_UP).doubleValue();
    }
}
