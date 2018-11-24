package com.word.poi.demo.util;

import com.word.poi.demo.pojo.PerformanceBriefInfo;
import org.apache.commons.lang3.StringUtils;

import java.text.DecimalFormat;


public class BriefHandler {

    /*根据名获取统计值*/
    public static String getValueByCity(PerformanceBriefInfo briefInfo, String city) {
        String statisticalValue = "";
        switch (city) {
            case "中国":
                statisticalValue = briefInfo.getStatistical_cz_value();
                break;
            case "俄罗斯":
                statisticalValue = briefInfo.getStatistical_dg_value();
                break;
            case "印度":
                statisticalValue = briefInfo.getStatistical_fs_value();
                break;
            case "美国":
                statisticalValue = briefInfo.getStatistical_gz_value();
                break;
            case "英国":
                statisticalValue = briefInfo.getStatistical_hy_value();
                break;
            case "德国":
                statisticalValue = briefInfo.getStatistical_hz_value();
                break;
            case "法国":
                statisticalValue = briefInfo.getStatistical_jm_value();
                break;
            case "日本":
                statisticalValue = briefInfo.getStatistical_jy_value();
                break;
            case "韩国":
                statisticalValue = briefInfo.getStatistical_mm_value();
                break;
            case "朝鲜":
                statisticalValue = briefInfo.getStatistical_mz_value();
                break;
            case "蒙古":
                statisticalValue = briefInfo.getStatistical_qy_value();
                break;
            case "巴西":
                statisticalValue = briefInfo.getStatistical_st_value();
                break;
            case "南非":
                statisticalValue = briefInfo.getStatistical_sw_value();
                break;
            case "尼泊尔":
                statisticalValue = briefInfo.getStatistical_sg_value();
                break;
            case "巴基斯坦":
                statisticalValue = briefInfo.getStatistical_sz_value();
                break;
            case "意大利":
                statisticalValue = briefInfo.getStatistical_yj_value();
                break;
            case "葡萄牙":
                statisticalValue = briefInfo.getStatistical_yf_value();
                break;
            case "比利时":
                statisticalValue = briefInfo.getStatistical_zj_value();
                break;
            case "菲律宾":
                statisticalValue = briefInfo.getStatistical_zq_value();
                break;
            case "迪拜":
                statisticalValue = briefInfo.getStatistical_zs_value();
                break;
            case "澳大利亚":
                statisticalValue = briefInfo.getStatistical_zh_value();
                break;
            case "全球":
                statisticalValue = briefInfo.getStatistical_all_value();
                break;
            case "一组":
                statisticalValue = briefInfo.getStatistical_yzz_value();
                break;
            case "二组":
                statisticalValue = briefInfo.getStatistical_ydnz_value();
                break;
            case "三组":
                statisticalValue = briefInfo.getStatistical_ydbz_value();
                break;
            case "四组":
                statisticalValue = briefInfo.getStatistical_yxybz_value();
                break;
        }
        return statisticalValue;
    }

    /**
     * 运算
     * @param aStr
     * @param bStr
     * @param flag  (a - b)/b :if(b==0||(a-b)==0) return "NULL";
     * @return
     *
     */
    public static String mathematics(String aStr, String bStr,String flag) {
        try {
            if (StringUtils.isEmpty(aStr)){
                return "NULL";
            }
            if ("null".equals(aStr)||"NULL".equals(aStr)){
                return "NULL";
            }
            if (StringUtils.isEmpty(bStr)){
                return "NULL";
            }
            if ("null".equals(bStr)||"NULL".equals(bStr)){
                return "NULL";
            }
            Double a = Double.parseDouble(aStr);
            Double  b = Double.parseDouble(bStr);
            DecimalFormat df = new DecimalFormat("0.0000");
            DecimalFormat dfTwo = new DecimalFormat("0.00");
            DecimalFormat dfInt = new DecimalFormat("0");
            if(flag.equals("(a - b)/b")){
                if(b==0||(a-b)==0) return "NULL";
                return df.format((a - b)/b);
            }else if(flag.equals("a-b")){
                return df.format(a-b);
            }else if(flag.equals("a/b")){//占比
                if(b==0) return "NULL";
                return df.format(a/b);
            }else if(flag.equals("aInt-bInt")){//数量对比 int类型
                return dfInt.format(a-b);
            }else if(flag.equals("a*b")){//数量对比 int类型
                return dfTwo.format(a*b);
            }else{
                return "NULL";
            }
        } catch (Exception e) {
            e.printStackTrace();
            return "NULL";
        }
    }


    private String formatData(String value){
        try {
            if (StringUtils.isNotEmpty(value)){
                if ("null".equals(value)){
                    return value;
                }
                Double cny = Double.parseDouble(value);//转换成Double
                DecimalFormat df = new DecimalFormat("#0.0000");//格式化
                return df.format(cny);
            }
            return value;
        } catch (NumberFormatException e) {
            e.printStackTrace();
            return value;
        }
    }
}
