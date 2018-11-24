package com.word.poi.demo.mapper;

import com.word.poi.demo.pojo.DataDemo;
import com.word.poi.demo.util.ExportToExcelUtils;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;


public class SimulatedDb {

    String [] PERFORMANCEBRIEFINFO = {"statistics_directory","statistical_term","statistical_all_value","statistical_cz_value","statistical_dg_value",
            "statistical_fs_value","statistical_gz_value","statistical_hy_value","statistical_hz_value","statistical_jy_value",
            "statistical_mm_value","statistical_mz_value","statistical_qy_value","statistical_st_value","statistical_sw_value",
            "statistical_sg_value","statistical_sz_value","statistical_yj_value","statistical_yf_value","statistical_zj_value",
            "statistical_zq_value","statistical_zs_value","statistical_zh_value","statistical_jm_value","id","statistical_yzz_value",
            "statistical_ydnz_value","statistical_ydbz_value","statistical_yxybz_value"};

    public  List<DataDemo> getSimulatedDb(){
        List<DataDemo> parseList = new ArrayList<>();
        try {
            ExportToExcelUtils<DataDemo> util = new ExportToExcelUtils<>();
            InputStream in = this.getClass().getResourceAsStream("/templates/simulatedDb.xlsx");
            parseList =  util.excelParseList(in, DataDemo.class,PERFORMANCEBRIEFINFO,0);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return parseList;
    }



}

