package com.word.poi.demo.service;

import com.word.poi.demo.mapper.SimulatedDb;
import com.word.poi.demo.pojo.PerformanceBriefInfo;
import com.word.poi.demo.util.BriefHandler;
import com.word.poi.demo.util.WorderToNewWordUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.util.*;

@Service
public class WordServiceImpl implements WordService {
    //顺序不能动，与word列表的行数对应
    String[] citys = {"全球","美国","俄罗斯","中国","英国","德国","法国","一组","韩国","日本","朝鲜","印度",
            "迪拜","二组","菲律宾","澳大利亚","加拿大","尼泊尔","三组","意大利","葡萄牙","比利时","南非","巴基斯坦","蒙古","四组"};
    String [] percents = {"C8","C9","C14","C15","C18","C19","C22","C23","C239","C25","C26","C244","C28","C29","C333","C336","C337",
            "C30C30S(a-b)/b","C31C31S(a-b)/b","C35","C33","C191C191S(a-b)/b","C192C192S(a-b)/b",
            "C193C193S(a-b)/b","C261","C194C194Sa-b","C195C195Sa-b","C196C196Sa-b","C197C197Sa-b","C198C198Sa-b","C199C199Sa-b",
            "C200C200Sa-b","C52S","C150S","C47S","C50S","C53S","C56S",
            "C59S","C62S","C65S","C37","C45","C254","C15C260a-b","C49","C52","C55","C58","C61","C64","C67",
            "C150C150Sa-b","C48","C51","C54","C57","C60","C63","C66",
            "C256S","C153S","C79S","C82S","C85S","C88S","C91S","C94S","C97S","C69","C77","C258","C153C260a-b","C81",
            "C84","C87","C90","C93","C96","C99","C74C74S(a-b)/b","C76C76S(a-b)/b","C257","C153C153Sa-b","C80","C83","C86",
            "C89","C92","C95","C98","C252","C47","C50","C53","C56","C59","C62","C65","C267","C205C38a/b","C111","C113",
            "C115","C117","C119","C121","C123","C262C38a/b","C106C36a/b","C108C42a/b","C110C44a/b","C267C252a-b","C205C38C150a/b-c",
            "C111C47a-b","C113C50a-b","C115C53a-b","C117C56a-b","C119C59a-b","C121C62a-b","C123C65a-b","C256","C79","C82",
            "C85","C88","C91","C94","C97","C271","C208C70a/b","C130","C132","C134","C136","C138","C140","C142","C263C70a/b",
            "C125C68a/b","C127C74a/b","C129C76a/b","C271C256a/b","C208C70C153a/b-c","C130C79a-b","C132C82a-b","C134C85a-b",
            "C136C88a-b","C138C91a-b","C140C94a-b","C142C97a-b","C147","C148","C150","C151","C153","C154","C156","C157","C159",
            "C160","C162","C163","C165","C166","C168","C169","C171","C172","C286","C290","C293","C294",
            "C298","C306","C310","C314","C318","C175","C175C175S(a-b)/b","C285","C289",
            "C297","C184","C184C184S(a-b)/b","C305","C309","C313","C317","C203","C206","C209","C212","C215","C218","C221",
            "C224","C227","C325","C326","C328","C329","C177","C179","C181","C186","C188","C190",
            "C250","C249","C260","C194","C195","C196","C197","C198","C199","C200",
            "C249S","C260S","C194S","C195S","C196S","C197S","C198S","C199S","C200S","C42C192a/b","C150C260a-b",
            "C38C38S(a-b)/b","C36C36S(a-b)/b","C42C42S(a-b)/b","C44C44S(a-b)/b","C74C192a/b","C70C70S(a-b)/b","C68C68S(a-b)/b",
            "C205C262a/b","C205C262C150a/b-c","C208C263a/b","C208C263C153a/b-c","C271C256a-b","C39","C71"

    };
    String [] plusSigns = {"C3C3Sa-b","C5","C7","C9","C13C13Sa-b","C15","C16C16Sa-b","C17C17Sa-b",
            "C21C21Sa-b","C23","C238","C24C24Sa-b","C243","C27C27Sa-b","C332","C335","C31C31S(a-b)/b",
            "C151","C154","C160","C163","C169","C172","C286","C294","C298","C306",
            "C314","C318","C12C12Sa-b","C17C17Sa-b","C19","C20C20Sa-b",
            "C26","C29","C337",
            "C290","C310","","","",""
    };

    //动态 增加和减少
    String [] upAndDown = {"C30C30S(a-b)/b","C148","C157","C166","C326","C329"};

    //正值为   劣化ＸＸ%
    //负值为： 改善ＸＸ%
    String [] badAndGood = {"C175C175S(a-b)/b","C184C184S(a-b)/b","","",""};

    //汇总
    String[] columnFortable0 ={"C12","C275","C276","C279%","C277","C280%","C278","C281%","C12+C278"};
    //20181017新增
    String[] columnFortable3 ={"C174","C17","C174/C17","C183","C21","C183/C21"};
    //表格 1
    String[] columnFortable4 ={"C3","C10+C13","C4","C6","C8%","C10","C11","C14%","C16","C17","C18%","C20","C21","C22%"};
    //表格 2    \\ 表示只是两个数相除 ，不求百分比
    String[] columnFortable5 ={"C24","C25%","C39%","C71%","C27","C28%","C31","C32","C33%","C34","C35%","C40","C72","C36\\C38","C68\\C70"};
    //表格 3
    String[] columnFortable6 ={"C38","C339%","C36","C41%","C42","C43%","C44","C46%","C252%","C253%","C150%","C151%","C47%","C48%"};
    //表格 4
    String[] columnFortable7 ={"C50%","C51%","C53%","C54%","C56%","C57%","C59%","C60%","C62%","C63%","C65%","C66%"};
    //表格 5
    String[] columnFortable8 ={"C70","C340%","C68","C73%","C74","C75%","C76","C78%","C256%","C257%","C153%","C154%","C79%","C80%"};
    //表格 6
    String[] columnFortable9 ={"C82%","C83%","C85%","C86%","C88%","C89%","C91%","C92%","C94%","C95%","C97%","C98%"};
    //表格 7
    String[] columnFortable10 ={"C103","C38","C105","C70","C106","C106/C36","C108","C108/C42","C110","C110/C193"};
    //表格 8
    String[] columnFortable11 ={"C111%","C112%","C113%","C114%","C115%","C116%","C117%","C118%","C119%","C120%","C121%","C122%","C123%","C124%"};
    //表格 9
    String[] columnFortable12 ={"C103","C38","C105","C70","C125","C125/C68","C127","C127/C74","C129","C129/C193"};
    //表格 10
    String[] columnFortable13 ={"C130%","C131%","C132%","C133%","C134%","C135%","C136%","C137%","C138%","C139%","C140%","C141%","C142%","C143%"};
    //表格 11
    String[] columnFortable14 ={"C146","C147%","C149","C150%","C152","C153%","C155","C156%","C158","C159%","C161","C162%","C164","C165%","C167","C168%","C170","C171%"};
    //表格 12
    String[] columnFortable15 ={"C284","C285%","C288","C289%","C292","C293%","C296","C297%","C304","C305%","C308","C309%","C312","C313%","C316","C317%"};
    //表格 13
    String[] columnFortable16 ={"C202","C203%","C205","C206%","C208","C209%","C211","C212%","C214","C215%","C217","C218%","C220","C221%","C223","C224%","C226","C227%"};
    //表格 14
    String[] columnFortable17 ={"C173","C174","C341","C342%","C175%","C176","C177%","C178","C179%","C180","C181%","C182","C183","C343","C344%","C184%","C185","C186%","C187","C188%","C189","C190%"};

    @Override
    public void exportWord(InputStream in,OutputStream out, String city) {
        try {
            SimulatedDb simulatedDb = new SimulatedDb();
            //本周
            List<PerformanceBriefInfo> thisFDDBriefs = simulatedDb.getSimulatedDb();
            Map<String, Map<String, String>> dataMap = convertCityResulMap(thisFDDBriefs, "C", "");
            Map<String, String> thisFDDBriefsMap = convertCityResulMap(thisFDDBriefs,"C","").get(city);
            //上周
            List<PerformanceBriefInfo> lastFDDBriefs = simulatedDb.getSimulatedDb();
            Map<String, String> lastFDDBriefsMap = convertCityResulMap(lastFDDBriefs,"C","S").get(city);
            //合并上周本周数据
            Map<String, String> idMap = new HashMap<>();
            idMap.putAll(thisFDDBriefsMap);
            idMap.putAll(lastFDDBriefsMap);
            //添加业务数据
            Map<String, String>  extendCidMap = extendDataForProvince(idMap);
            //添加各个全球排序数据
            extendCidMap = extendCityOrderData(extendCidMap,dataMap);
            //添加表格数据
            Map<String, List<List<String>>> extendTableMap = extendTableDataForProvince(dataMap);
            //根据业务逻辑添加表格颜色
            Map<String, List<String>>  colorMapForCell = setColorForCell(dataMap,extendCidMap);
            XWPFDocument docx = WorderToNewWordUtil.changWordForProvince(in, extendCidMap, extendTableMap,colorMapForCell);
            docx.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            try {
                if (out!=null)out.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }


    /**
     * 添加业务数据
     * @param idMap
     * @return
     */
    private Map<String, String> extendDataForProvince(Map<String, String> idMap) {
        idMap.put("startMM","10");
        idMap.put("startDD","09");
        idMap.put("endMM","10");
        idMap.put("endDD","19");
        idMap.put("deadline","10-19");


        idMap.put("C3C3Sa-b",this.compute(idMap.get("C3"),idMap.get("C3S"),"a-b"));
        idMap.put("C10C13a+b",this.compute(idMap.get("C10"),idMap.get("C13"),"a+b"));
        idMap.put("C12C12Sa-b",this.compute(idMap.get("C12"),idMap.get("C12S"),"a-b"));
        idMap.put("C13C13Sa-b",this.compute(idMap.get("C13"),idMap.get("C13S"),"a-b"));
        idMap.put("C14C15a+b",this.compute(idMap.get("C14"),idMap.get("C15"),"a+b"));
        idMap.put("C16C16Sa-b",this.compute(idMap.get("C16"),idMap.get("C16S"),"a-b"));
        idMap.put("C17C17Sa-b",this.compute(idMap.get("C17"),idMap.get("C17"),"a-b"));
        idMap.put("C20C20Sa-b",this.compute(idMap.get("C20"),idMap.get("C20S"),"a-b"));
        idMap.put("C21C21Sa-b",this.compute(idMap.get("C21"),idMap.get("C21S"),"a-b"));
        idMap.put("C24C24Sa-b",this.compute(idMap.get("C24"),idMap.get("C24S"),"a-b"));
        idMap.put("C27C27Sa-b",this.compute(idMap.get("C27"),idMap.get("C27S"),"a-b"));
        idMap.put("C30C30S(a-b)/b",this.compute(idMap.get("C30"),idMap.get("C30S"),"(a-b)/b"));
        idMap.put("C31C31S(a-b)/b",this.compute(idMap.get("C31"),idMap.get("C31S"),"(a-b)/b"));
        idMap.put("C191C191S(a-b)/b",this.compute(idMap.get("C191"),idMap.get("C191S"),"(a-b)/b"));
        idMap.put("C192C192S(a-b)/b",this.compute(idMap.get("C192"),idMap.get("C192S"),"(a-b)/b"));
        idMap.put("C193C193S(a-b)/b",this.compute(idMap.get("C193"),idMap.get("C193S"),"(a-b)/b"));
        idMap.put("C194C194Sa-b",this.compute(idMap.get("C194"),idMap.get("C194S"),"a-b"));
        idMap.put("C195C195Sa-b",this.compute(idMap.get("C195"),idMap.get("C195S"),"a-b"));
        idMap.put("C196C196Sa-b",this.compute(idMap.get("C196"),idMap.get("C196S"),"a-b"));
        idMap.put("C197C197Sa-b",this.compute(idMap.get("C197"),idMap.get("C197S"),"a-b"));
        idMap.put("C198C198Sa-b",this.compute(idMap.get("C198"),idMap.get("C198S"),"a-b"));
        idMap.put("C199C199Sa-b",this.compute(idMap.get("C199"),idMap.get("C199S"),"a-b"));
        idMap.put("C200C200Sa-b",this.compute(idMap.get("C200"),idMap.get("C200S"),"a-b"));
        idMap.put("C42C192a/b",this.compute(idMap.get("C42"),idMap.get("C192"),"a/b"));
        idMap.put("C150C260a-b",this.compute(idMap.get("C150"),idMap.get("C260"),"a-b"));
        idMap.put("C38C38S(a-b)/b",this.compute(idMap.get("C38"),idMap.get("C38S"),"(a-b)/b"));
        idMap.put("C36C36S(a-b)/b",this.compute(idMap.get("C36"),idMap.get("C36S"),"(a-b)/b"));
        idMap.put("C42C42S(a-b)/b",this.compute(idMap.get("C42"),idMap.get("C42S"),"(a-b)/b"));
        idMap.put("C44C44S(a-b)/b",this.compute(idMap.get("C44"),idMap.get("C44S"),"(a-b)/b"));
        idMap.put("C150C150Sa-b",this.compute(idMap.get("C150"),idMap.get("C150S"),"a-b"));
        idMap.put("C74C192a/b",this.compute(idMap.get("C74"),idMap.get("C192"),"a/b"));
        idMap.put("C153C260a-b",this.compute(idMap.get("C153"),idMap.get("C260"),"a-b"));
        idMap.put("C70C70S(a-b)/b",this.compute(idMap.get("C70"),idMap.get("C70S"),"(a-b)/b"));
        idMap.put("C68C68S(a-b)/b",this.compute(idMap.get("C68"),idMap.get("C68S"),"(a-b)/b"));
        idMap.put("C74C74S(a-b)/b",this.compute(idMap.get("C74"),idMap.get("C74S"),"(a-b)/b"));
        idMap.put("C76C76S(a-b)/b",this.compute(idMap.get("C76"),idMap.get("C76S"),"(a-b)/b"));
        idMap.put("C153C153Sa-b",this.compute(idMap.get("C153"),idMap.get("C153S"),"a-b"));
        idMap.put("C205C262a/b",this.compute(idMap.get("C205"),idMap.get("C262"),"a/b"));
        idMap.put("C262C38a/b",this.compute(idMap.get("C262"),idMap.get("C38"),"a/b"));
        idMap.put("C106C36a/b",this.compute(idMap.get("C106"),idMap.get("C36"),"a/b"));
        idMap.put("C108C42a/b",this.compute(idMap.get("C108"),idMap.get("C42"),"a/b"));
        idMap.put("C110C44a/b",this.compute(idMap.get("C110"),idMap.get("C44"),"a/b"));
        idMap.put("C267C252a-b",this.compute(idMap.get("C267"),idMap.get("C252"),"a-b"));
        idMap.put("C205C262C150a/b-c",this.compute(idMap.get("C205"),idMap.get("C262"),idMap.get("C150"),"a/b-c"));
        idMap.put("C111C47a-b",this.compute(idMap.get("C111"),idMap.get("C47"),"a-b"));
        idMap.put("C113C50a-b",this.compute(idMap.get("C113"),idMap.get("C50"),"a-b"));
        idMap.put("C115C53a-b",this.compute(idMap.get("C115"),idMap.get("C53"),"a-b"));
        idMap.put("C117C56a-b",this.compute(idMap.get("C117"),idMap.get("C56"),"a-b"));
        idMap.put("C119C59a-b",this.compute(idMap.get("C119"),idMap.get("C59"),"a-b"));
        idMap.put("C121C62a-b",this.compute(idMap.get("C121"),idMap.get("C62"),"a-b"));
        idMap.put("C123C65a-b",this.compute(idMap.get("C123"),idMap.get("C65"),"a-b"));
        idMap.put("C208C263a/b",this.compute(idMap.get("C208"),idMap.get("C263"),"a/b"));
        idMap.put("C263C70a/b",this.compute(idMap.get("C263"),idMap.get("C70"),"a/b"));
        idMap.put("C125C68a/b",this.compute(idMap.get("C125"),idMap.get("C68"),"a/b"));
        idMap.put("C127C74a/b",this.compute(idMap.get("C127"),idMap.get("C74"),"a/b"));
        idMap.put("C129C76a/b",this.compute(idMap.get("C129"),idMap.get("C76"),"a/b"));
        idMap.put("C271C256a-b",this.compute(idMap.get("C271"),idMap.get("C256"),"a-b"));
        idMap.put("C208C263C153a/b-c",this.compute(idMap.get("C208"),idMap.get("C263"),idMap.get("C153"),"a/b-c"));
        idMap.put("C130C79a-b",this.compute(idMap.get("C130"),idMap.get("C79"),"a-b"));
        idMap.put("C132C82a-b",this.compute(idMap.get("C132"),idMap.get("C82"),"a-b"));
        idMap.put("C134C85a-b",this.compute(idMap.get("C134"),idMap.get("C85"),"a-b"));
        idMap.put("C136C88a-b",this.compute(idMap.get("C136"),idMap.get("C88"),"a-b"));
        idMap.put("C138C91a-b",this.compute(idMap.get("C138"),idMap.get("C91"),"a-b"));
        idMap.put("C140C94a-b",this.compute(idMap.get("C140"),idMap.get("C94"),"a-b"));
        idMap.put("C142C97a-b",this.compute(idMap.get("C142"),idMap.get("C97"),"a-b"));
        idMap.put("C175C175S(a-b)/b",this.compute(idMap.get("C175"),idMap.get("C175S"),"(a-b)/b"));
        idMap.put("C184C184S(a-b)/b",this.compute(idMap.get("C184"),idMap.get("C184S"),"(a-b)/b"));
        idMap.put("C36C38a/b",this.compute(idMap.get("C36"),idMap.get("C38"),"a/b"));
        idMap.put("C68C70a/b",this.compute(idMap.get("C68"),idMap.get("C70"),"a/b"));

        //添加百分号
        for (String key : percents){
            Double a;
            try{a = Double.parseDouble(idMap.get(key));}catch (Exception e){ continue;}
            DecimalFormat df = new DecimalFormat("0.00%");
            idMap.put(key,df.format(a));
        }
        //添加+号
        for (String key : plusSigns){
            Double a;
            try{
                if (idMap.get(key).contains("%")){
                    a = Double.parseDouble(idMap.get(key).replace("%",""))*0.01;
                }else{
                    a = Double.parseDouble(idMap.get(key));
                }
            }catch (Exception e){ continue;}
            if (a > 0){
                idMap.put(key,"+"+idMap.get(key));
            }
        }
        //动态添加 增加和减少
        for (String key : upAndDown){
            Double a;
            try{
                if (idMap.get(key).contains("%")){
                    a = Double.parseDouble(idMap.get(key).replace("%",""))*0.01;
                }else{
                    a = Double.parseDouble(idMap.get(key));
                }
            }catch (Exception e){ continue;}
            if (a >= 0){
                idMap.put(key,"增加"+idMap.get(key));
            }else{
                idMap.put(key,"减少"+idMap.get(key).replace("-",""));
            }
        }
        //正值为   劣化ＸＸ%
        //负值为： 改善ＸＸ%
        for (String key : badAndGood){
            Double a;
            try{
                if (idMap.get(key).contains("%")){
                    a = Double.parseDouble(idMap.get(key).replace("%",""))*0.01;
                }else{
                    a = Double.parseDouble(idMap.get(key));
                }
            }catch (Exception e){ continue;}
            if (a >= 0){
                idMap.put(key,"劣化"+idMap.get(key));
            }else{
                idMap.put(key,"改善"+idMap.get(key).replace("-",""));
            }
        }
        return idMap;
    }

    /**
     * word 段落 top3等排序及业务数据添加
     * @param extendCidMap
     * @param dataMap
     * @return
     */
    private Map<String,String> extendCityOrderData(Map<String, String> extendCidMap, Map<String, Map<String, String>> dataMap) {
        //获取最高的三个数据加百分号
        String[] cityOrderTop3percents ={"C8","C14","C147","C156","C165","C175","C184","C325","C328","C281"};
        //获取最低的三个数据加百分号
        String[] cityOrderBottom3percents ={"C14"};
        //获取最低的三个数据加百分号并拼接上周的数据 用“（”隔开
        String[] cityOrderBottom3percentsAndLastWeek ={"C25(C26","C28(C29","C336(C337"};

        //获取最高的三个数据
        String[] cityOrderTop3 ={};
        //获取最低的三个数据
        String[] cityOrderBottom3 ={"C40","C72"};

        //统计为0 的数据
        String[] cityDataIsZero ={"C20"};

        //统计 优化劣化（数据最大的）的数据  FDD1800差小区占比环比上周，FDD900差小区占比环比上周
        String[] cityDataOrderMax ={"C175(a-b)/bC175S","C184(a-b)/bC184S"};//${C175(a-b)/bC175SMax}   ${C184(a-b)/bC184SMax}

        for (String cNo:cityOrderTop3percents) {
            String top3Str = getTop3FromMapByCNo(dataMap,cNo,"%");
            extendCidMap.put(cNo+"Top3",top3Str);
        }
        for (String cNo:cityOrderBottom3percents) {
            String bottom3Str = getBottom3FromMapByCNo(dataMap,cNo,"%");
            extendCidMap.put(cNo+"Bottom3",bottom3Str);
        }
        for (String cNo:cityOrderTop3) {
            String top3Str = getTop3FromMapByCNo(dataMap,cNo,"");
            extendCidMap.put(cNo+"Top3",top3Str);
        }
        for (String cNo:cityOrderBottom3) {
            String bottom3Str = getBottom3FromMapByCNo(dataMap,cNo,"");
            extendCidMap.put(cNo+"Bottom3",bottom3Str);
        }
        for (String cNo:cityDataIsZero) {
            String cityZeroStr = getCityIsZeroFromMapByCNo(dataMap,cNo);
            extendCidMap.put(cNo+"Zero",cityZeroStr);
        }

        for (String cNo:cityOrderBottom3percentsAndLastWeek) {
            String bottom3Str = getCityOrderBottom3percentsAndLastWeek(dataMap,cNo,"%");
            extendCidMap.put(cNo+"Bottom3",bottom3Str);
        }
        for (String cNo:cityDataOrderMax) {
            String bottom3Str = getCityOrderMax(dataMap,cNo,"%");
            extendCidMap.put(cNo+"Max",bottom3Str);
        }
        return extendCidMap;
    }

    private String getCityOrderMax(Map<String, Map<String, String>> dataMap, String cNo, String flag) {
        //${C175(a-b)/bC175SMax}
        DecimalFormat df = new DecimalFormat("0.00%");
        String[] split = cNo.split("\\(a-b\\)/b");
        String cb = split[0];//本周
        String csb = split[1];//本周比上周
        Map<String, Double> C175C175S = getCityMapDataBycNo(dataMap,cb,csb,"(a-b)/b");
        C175C175S = sortByValueDescending(C175C175S);//降序排序（从高到低）
        Map.Entry<String, Double> max = C175C175S.entrySet().iterator().next();
        Double value = max.getValue();
        if ("%".equals(flag)){
            String format = df.format(value);
            if(value>0){
                format = "+"+format;
            }
            return max.getKey()+"差小区占比环比上周"+format;
        }
        return "";
    }


    /**
     * 添加 本周比上周的数据
     * @param dataMap
     * @param cNo
     * @param flag
     * @return
     */
    private String getCityOrderBottom3percentsAndLastWeek(Map<String, Map<String, String>> dataMap, String cNo, String flag) {
        String[] split = cNo.split("\\(");
        String cb = split[0];//本周
        String csb = split[1];//本周比上周
        Map<String, Double> bottom3 = getCityMapDataBycNo(dataMap, cb);
        bottom3 = sortByValueAscending(bottom3);//升序排序（从低到高）
        StringBuilder front3sb = new StringBuilder();
        DecimalFormat df = new DecimalFormat("0.00%");
        int count = 0;
        for (Map.Entry<String, Double> entry : bottom3.entrySet()) {
            if (count <3){
                if ("%".equals(flag)){
                    String csbStr = dataMap.get(entry.getKey()).get(csb);
                    if (StringUtils.isEmpty(csbStr)||"NULL".equals(csbStr)){
                        csbStr = "0";
                    }
                    double csbDoub = Double.parseDouble(csbStr);
                    String format = df.format(csbDoub);//加百分号
                    if (csbDoub > 0){
                        format = "+"+format;
                    }
                    front3sb.append(entry.getKey()+ df.format(entry.getValue())+"("+format+")"+"、");
                }else {
                    String csbStr = dataMap.get(entry.getKey()).get(csb);
                    if (StringUtils.isEmpty(csbStr)||"NULL".equals(csbStr)){
                        csbStr = "0";
                    }
                    double csbDoub = Double.parseDouble(csbStr);
                    String format = df.format(csbDoub);//加百分号
                    if (csbDoub > 0){
                        format = "+"+format;
                    }
                    front3sb.append(entry.getKey()+ entry.getValue()+"("+format+")"+"、");
                }
            }
            count++;
        }
        return  front3sb.toString().substring(0,front3sb.toString().length()-1);
    }

    private String getCityIsZeroFromMapByCNo(Map<String, Map<String, String>> dataMap, String cNo) {
        Map<String, Double> top3 = getCityMapDataBycNo(dataMap, cNo);
        StringBuilder cityDataIsZero = new StringBuilder();

        for (Map.Entry<String, Double> entry : top3.entrySet()) {
            if (entry.getValue()==0)
                cityDataIsZero.append(entry.getKey()+"、");
        }
        if (cityDataIsZero.length()>0){
            return cityDataIsZero.toString().substring(0,cityDataIsZero.toString().length()-1);
        }
        return cityDataIsZero.toString();
    }

    /**
     * 通过C编号获取全部的 全球数据，再根据全球数据排序，返回最高的三个全球数据
     * @param dataMap 所有的数据
     * @param cNo  C编号
     * @param flag 判断返回的数据 是否要加% 等符号
     * @return
     */
    private String getTop3FromMapByCNo(Map<String, Map<String, String>> dataMap,String cNo,String flag) {
        Map<String, Double> top3 = getCityMapDataBycNo(dataMap, cNo);
        top3 = sortByValueDescending(top3);//降序排序（从高到低）
        return getFront3FromMap(top3,flag);
    }


    /**
     * 通过C编号获取全部的 全球数据，再根据全球数据排序，返回最高的三个全球数据
     * @param dataMap 所有的数据
     * @param cNo  C编号
     * @param flag 判断返回的数据 是否要加% 等符号
     * @return
     */
    private String getBottom3FromMapByCNo(Map<String, Map<String, String>> dataMap,String cNo,String flag) {
        Map<String, Double> bottom3 = getCityMapDataBycNo(dataMap, cNo);
        bottom3 = sortByValueAscending(bottom3);//升序排序（从低到高）
        return getFront3FromMap(bottom3,flag);
    }

    private Map<String, Double> getCityMapDataBycNo(Map<String, Map<String, String>> dataMap, String cNo) {
        String[] onlyCitys = {"美国","俄罗斯","中国","英国","德国","法国","韩国","日本","朝鲜","印度",
                "迪拜","菲律宾","澳大利亚","加拿大","尼泊尔","意大利","葡萄牙","比利时","南非","巴基斯坦","蒙古"};
        Map<String, Double> map = new HashMap<>();
        for (String city:onlyCitys) {
            Map<String, String> cityMap = dataMap.get(city);
            String cNoValueStr;
            if (cNo.contains("/")){
                String[] split = cNo.split("/");
                String cellStr1 = cityMap.get(split[0]);
                String cellStr2 = cityMap.get(split[1]);
                cNoValueStr = compute(cellStr1, cellStr2, "a/b");
            }else if(cNo.contains("+")){
                String[] split = cNo.split("\\+");
                String cellStr1 = cityMap.get(split[0]);
                String cellStr2 = cityMap.get(split[1]);
                cNoValueStr = compute(cellStr1, cellStr2, "a+b");
            }else if(cNo.contains("\\")){
                String[] split = cNo.split("\\\\");
                String cellStr1 = cityMap.get(split[0]);
                String cellStr2 = cityMap.get(split[1]);
                cNoValueStr = compute(cellStr1, cellStr2, "a/b");
            } else {
                cNoValueStr = cityMap.get(cNo);
            }
            if (StringUtils.isEmpty(cNoValueStr)||"NULL".equals(cNoValueStr)){
                cNoValueStr = "0";
            }
            Double dub =  Double.parseDouble(cNoValueStr);
            map.put(city,dub);
        }
        return map;
    }

    /**
     * 全球作为key ，把各个全球计算的数据put进map
     * @param dataMap
     * @param cb
     * @param csb
     * @param flag
     * @return
     */
    private Map<String,Double> getCityMapDataBycNo(Map<String, Map<String, String>> dataMap, String cb, String csb, String flag) {
        String[] onlyCitys = {"美国","俄罗斯","中国","英国","德国","法国","韩国","日本","朝鲜","印度",
                "迪拜","菲律宾","澳大利亚","加拿大","尼泊尔","意大利","葡萄牙","比利时","南非","巴基斯坦","蒙古"};
        Map<String, Double> map = new HashMap<>();
        for (String city:onlyCitys) {
            Map<String, String> cityMap = dataMap.get(city);
            String compute = compute(cityMap.get(cb), cityMap.get(csb), flag);
            if (StringUtils.isEmpty(compute)||"NULL".equals(compute)){
                compute = "0";
            }
            Double dub =  Double.parseDouble(compute);
            map.put(city,dub);
        }
        return map;
    }

    /**
     * 获取map 的前三个数据
     * @param top3  全球数据
     * @param flag  判断返回的数据 是否要加% 等符号
     * @return
     */
    private String getFront3FromMap(Map<String, Double> top3, String flag) {
        StringBuilder front3sb = new StringBuilder();
        DecimalFormat df = new DecimalFormat("0.00%");
        int count = 0;
        for (Map.Entry<String, Double> entry : top3.entrySet()) {
            if (count <3){
                if ("%".equals(flag)){
                    front3sb.append(entry.getKey()+ df.format(entry.getValue())+"、");
                }else {
                    front3sb.append(entry.getKey() +"("+entry.getValue()+")、");
                }
            }
            count++;
        }
        return  front3sb.toString().substring(0,front3sb.toString().length()-1);
    }

    /**
     * 降序排序 (从大到小)
     * @param map
     * @param <K>
     * @param <V>
     * @return
     */
    public  <K, V extends Comparable<? super V>> Map<K, V>  sortByValueDescending(Map<K, V> map) {
        List<Map.Entry<K, V>> list = new LinkedList<>(map.entrySet());
        Collections.sort(list, new Comparator<Map.Entry<K, V>>() {
            @Override
            public int compare(Map.Entry<K, V> o1, Map.Entry<K, V> o2)
            {
                int compare = (o1.getValue()).compareTo(o2.getValue());
                return -compare;
            }
        });

        Map<K, V> result = new LinkedHashMap<K, V>();
        for (Map.Entry<K, V> entry : list) {
            result.put(entry.getKey(), entry.getValue());
        }
        return result;
    }

    /**
     * 升序排序（从小到大）
     * @param map
     * @param <K>
     * @param <V>
     * @return
     */
    public  <K, V extends Comparable<? super V>> Map<K, V> sortByValueAscending(Map<K, V> map) {
        List<Map.Entry<K, V>> list = new LinkedList<Map.Entry<K, V>>(map.entrySet());
        Collections.sort(list, new Comparator<Map.Entry<K, V>>()
        {
            @Override
            public int compare(Map.Entry<K, V> o1, Map.Entry<K, V> o2)
            {
                int compare = (o1.getValue()).compareTo(o2.getValue());
                return compare;
            }
        });

        Map<K, V> result = new LinkedHashMap<K, V>();
        for (Map.Entry<K, V> entry : list) {
            result.put(entry.getKey(), entry.getValue());
        }
        return result;
    }

    /**
     * 添加表格数据
     * @param dataMap
     * @return
     */
    private Map<String, List<List<String>>> extendTableDataForProvince(Map<String, Map<String, String>> dataMap) {
        Map<String, List<List<String>>> tableMap = new HashMap<>();
        List<List<String>> table0 = getTableByCid(columnFortable0,dataMap);
        List<List<String>> table3 = getTableByCid(columnFortable3,dataMap);
        List<List<String>> table4 = getTableByCid(columnFortable4,dataMap);
        List<List<String>> table5 = getTableByCid(columnFortable5,dataMap);
        List<List<String>> table6 = getTableByCid(columnFortable6,dataMap);
        List<List<String>> table7 = getTableByCid(columnFortable7,dataMap);
        List<List<String>> table8 = getTableByCid(columnFortable8,dataMap);
        List<List<String>> table9 = getTableByCid(columnFortable9,dataMap);
        List<List<String>> table10 = getTableByCid(columnFortable10,dataMap);
        List<List<String>> table11 = getTableByCid(columnFortable11,dataMap);
        List<List<String>> table12 = getTableByCid(columnFortable12,dataMap);
        List<List<String>> table13 = getTableByCid(columnFortable13,dataMap);
        List<List<String>> table14 = getTableByCid(columnFortable14,dataMap);
        List<List<String>> table15 = getTableByCid(columnFortable15,dataMap);
        List<List<String>> table16 = getTableByCid(columnFortable16,dataMap);
        List<List<String>> table17 = getTableByCid(columnFortable17,dataMap);
        tableMap.put("table0",table0 );
        tableMap.put("table3",table3 );
        tableMap.put("table4",table4 );
        tableMap.put("table5",table5 );
        tableMap.put("table6",table6 );
        tableMap.put("table7",table7 );
        tableMap.put("table8",table8 );
        tableMap.put("table9",table9 );
        tableMap.put("table10",table10);
        tableMap.put("table11",table11);
        tableMap.put("table12",table12);
        tableMap.put("table13",table13);
        tableMap.put("table14",table14);
        tableMap.put("table15",table15);
        tableMap.put("table16",table16);
        tableMap.put("table17",table17);
        return tableMap;
    }

    private Map<String, List<String>> setColorForCell(Map<String, Map<String, String>> dataMap, Map<String, String> extendCidMap) {
        //汇总
        String[] colorFortableTop0 = {"C280%", "C281%"};//标红最大的3个值
        String[] colorFortableBottom0 = {"C279%"};//标红最大的3个值
        //需要标红的数据和规则
        Map<String,String> colorFortableTop1 = table1AddColorData();//列表固定坐标 根据门限 添加颜色
        Map<String,String> colorFortableTop2 = table2AddColorData();//列表固定坐标 根据门限 添加颜色
        //20181017新增列表
        String[] colorFortableTop3 = {"C174","C174/C17","C183","C183/C21"};
        //表格 1
        String[] colorFortableTop4 = {"C8%"};
        //表格 2
        String[] colorFortableBottom5 = {"C25%","C28%","C40","C72","C36\\C38","C68\\C70"};//标红最小的3个值
        //表格 3
        String[] colorFortableTop6 = {"C252%","C150%"};//标红最大的3个值
        String[] colorFortableBottom6 = {"C47%"};//标红最小的3个值
        String[] colorFortableTopPositive6 = {"C253%","C151%","C48%"};//标红最大且大于零的3个值
        //表格 4
        String[] colorFortableTop7 = {"C50%","C51%","C59%","C60%","C65%","C66%"};
        String[] colorFortableBottom7 = {"C53%","C54%","C56%","C57%","C62%","C63%"};
        //表格 5全网FDD900性能
        String[] colorFortableTop8 = {"C256%","C257%","C153%","C154%"};
        String[] colorFortableBottom8 = {"C79%","C80%"};
        //表格 6
        String[] colorFortableTop9 = {"C82%","C83%","C91%","C92%","C97%","C98%"};
        String[] colorFortableBottom9 = {"C85%","C86%","C88%","C89%","C94%","C95%"};
        //表格 7
        //String[] colorFortable10 ={"C103","C38","C105","C70","C106","C106/C36","C108","C108/C42","C110","C110/C193"};
        //表格 8
        String[] colorFortableTop11 = {"C113%","C114%","C119%","C120%","C123%","C124%"};
        String[] colorFortableBottom11 = {"C111%","C112%","C115%","C116%","C117%","C118%","C121%","C122%"};
        //表格 9
        //String[] colorFortable12 ={"C103","C38","C105","C70","C125","C126","C127","C128","C129","C129/C193"};
        //表格 10
        String[] colorFortableTop13 = {"C132%","C133%","C138%","C139%","C142%","C143%"};
        String[] colorFortableBottom13 = {"C130%","C131%","C134%","C135%","C136%","C137%","C140%","C141%"};
        //表格 11
        String[] colorFortableTop14 = columnFortable14;
        //表格 12
        String[] colorFortableTop15 = columnFortable15;
        //表格 13
        String[] colorFortableTop16 = columnFortable16;
        //表格 14
        String[] colorFortableTop17 = {"C174","C175%","C183","C184%"};

        Map<String, List<String>> colorMapForCell = new HashMap<>();
        List<String> table0 = setColorTableForCell(colorFortableTop0,columnFortable0,dataMap,"Top3");//给table2设置颜色
        table0.addAll(setColorTableForCell(colorFortableBottom0,columnFortable0,dataMap,"Bottom3"));//给table2设置颜色

        List<String> table1 = setColorTableForCell(extendCidMap,colorFortableTop1);//给table1设置颜色
        List<String> table2 = setColorTableForCell(extendCidMap,colorFortableTop2);//给table2设置颜色

        List<String> table3 = setColorTableForCell(colorFortableTop3,columnFortable3,dataMap, "Top3");//给table2设置颜色
        List<String> table4 = setColorTableForCell(colorFortableTop4,columnFortable4,dataMap, "Top3");//给table2设置颜色
        List<String> table5 = setColorTableForCell(colorFortableBottom5,columnFortable5,dataMap, "Bottom3");//给table2设置颜色
        List<String> table6 = setColorTableForCell(colorFortableTop6,columnFortable6,dataMap, "Top3");//给table2设置颜色
        table6.addAll(setColorTableForCellTopPositive(colorFortableTopPositive6,columnFortable6,dataMap));//标红最大且大于零的3个值
        table6.addAll(setColorTableForCell(colorFortableBottom6,columnFortable6,dataMap,"Bottom3"));//标红最小3个值

        List<String> table7 = setColorTableForCell(colorFortableTop7,columnFortable7,dataMap, "Top3");//给table2设置颜色
        table7.addAll(setColorTableForCell(colorFortableBottom7,columnFortable7,dataMap, "Bottom3"));//标红最小的3个值

        List<String> table8 = setColorTableForCell(colorFortableTop8,columnFortable8,dataMap, "Top3");//给table2设置颜色
        table8.addAll(setColorTableForCell(colorFortableBottom8,columnFortable8,dataMap, "Bottom3"));//标红最小的3个值

        List<String> table9 = setColorTableForCell(colorFortableTop9,columnFortable9,dataMap, "Top3");//给table2设置颜色
        table9.addAll(setColorTableForCell(colorFortableBottom9,columnFortable9,dataMap, "Bottom3"));//标红最小的3个值

        List<String> table11 = setColorTableForCell(colorFortableTop11,columnFortable11,dataMap, "Top3");//给table2设置颜色
        table11.addAll(setColorTableForCell(colorFortableBottom11,columnFortable11,dataMap, "Bottom3"));//标红最小的3个值

        List<String> table13 = setColorTableForCell(colorFortableTop13,columnFortable13,dataMap, "Top3");//给table2设置颜色
        table13.addAll(setColorTableForCell(colorFortableBottom13,columnFortable13,dataMap, "Bottom3"));//标红最小的3个值

        List<String> table14 = setColorTableForCell(colorFortableTop14,columnFortable14,dataMap, "Top3");//给table2设置颜色
        List<String> table15 = setColorTableForCell(colorFortableTop15,columnFortable15,dataMap, "Top3");//给table2设置颜色
        List<String> table16 = setColorTableForCell(colorFortableTop16,columnFortable16,dataMap, "Top3");//给table2设置颜色
        List<String> table17 = setColorTableForCell(colorFortableTop17,columnFortable17,dataMap, "Top3");//给table16设置颜色

        colorMapForCell.put("table0",table0);
        colorMapForCell.put("table1",table1);
        colorMapForCell.put("table2",table2);
        colorMapForCell.put("table3",table3);
        colorMapForCell.put("table4", table4);
        colorMapForCell.put("table5", table5);
        colorMapForCell.put("table6", table6);
        colorMapForCell.put("table7", table7);
        colorMapForCell.put("table8", table8);
        colorMapForCell.put("table9", table9);
        colorMapForCell.put("table11", table11);
        colorMapForCell.put("table13", table13);
        colorMapForCell.put("table14", table14);
        colorMapForCell.put("table15", table15);
        colorMapForCell.put("table16", table16);
        colorMapForCell.put("table17", table17);
        return colorMapForCell;
    }

    private List<String> setColorTableForCell(Map<String, String> extendCidMap, Map<String, String> colorFortableTop1) {
        List<String> table = new ArrayList<>();//给table设置颜色
        Set<Map.Entry<String, String>> entries = colorFortableTop1.entrySet();
        for (Map.Entry<String, String> entry:entries) {
            String key = entry.getKey();//坐标
            String value = entry.getValue();
            String[] split = value.split(",");
            String cNo = split[0];
            String flag = split[1];
            if (alterColor(extendCidMap.get(cNo),flag)){
                table.add(key);
            }
        }

        return table;
    }

    /**
     * 标红最大或者最小的3个值
     * @param colorFortable
     * @param columnFortable
     * @param dataMap
     * @param flag :Top3 标红最大的3个值, Bottom3 标红最小的3个值
     * @return
     */
    public List<String> setColorTableForCell(String[] colorFortable, String[] columnFortable,
                                             Map<String, Map<String, String>> dataMap, String flag) {
        List<String> table = new ArrayList<>();//给table设置颜色
        for (String cNoForColor: colorFortable) {
            String columnNoByCNo = getColumnNoByCNo(cNoForColor, columnFortable);//得到table 的列编号
            if(cNoForColor.contains("%")){
                cNoForColor = cNoForColor.split("%")[0];//去除百分号
            }
            int count = 0;
            Map<String, Double> c8 = getCityMapDataBycNo(dataMap, cNoForColor);
            Map<String, Double> sortMap =  new LinkedHashMap();
            if ("Top3".equals(flag)){
                sortMap = sortByValueDescending(c8);//降序排序 (从大到小)，获取最大的三位
            }else if("Bottom3".equals(flag)){
                sortMap = sortByValueAscending(c8);//升序排序（从小到大），获取最小的三位
            }
            for (Map.Entry<String, Double> entry : sortMap.entrySet()) {
                if (count < 3) {
                    String rowNoByCity = getRowNoByCity(entry.getKey());
                    table.add(rowNoByCity + "," + columnNoByCNo);
                }
                count++;
            }
        }
        return table;
    }

    /**
     * 标红最大且大于零的3个值
     * @param colorFortable
     * @param columnFortable
     * @param dataMap
     * @return
     */
    public List<String> setColorTableForCellTopPositive(String[] colorFortable, String[] columnFortable,
                                                        Map<String, Map<String, String>> dataMap) {
        List<String> table = new ArrayList<>();//给table设置颜色
        for (String cNoForColor: colorFortable) {
            String columnNoByCNo = getColumnNoByCNo(cNoForColor, columnFortable);//得到table 的列编号
            if(cNoForColor.contains("%")){
                cNoForColor = cNoForColor.split("%")[0];//去除百分号
            }
            int count = 0;
            Map<String, Double> c8 = getCityMapDataBycNo(dataMap, cNoForColor);
            Map<String, Double> top3 = sortByValueDescending(c8);//降序排序 (从大到小)，获取最大的三位
            for (Map.Entry<String, Double> entry : top3.entrySet()) {
                if (count < 3) {
                    if (entry.getValue()>0){
                        String rowNoByCity = getRowNoByCity(entry.getKey());
                        table.add(rowNoByCity + "," + columnNoByCNo);
                    }
                }
                count++;
            }
        }
        return table;
    }

    private List<List<String>> getTableByCid(String[] columnFortableNo, Map<String, Map<String, String>> dataMap) {
        List<List<String>> tableList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00%");
        for (String city:citys) {
            Map<String, String> cityMap = dataMap.get(city);
            List<String> rowList = new ArrayList<>();
            rowList.add(city);//添加全球
            for (String cid:columnFortableNo) {
                if (cid.contains("+")){
                    String[] split = cid.split("\\+");
                    String cellStr1 = cityMap.get(split[0]);
                    String cellStr2 = cityMap.get(split[1]);
                    //int cell1 = Integer.parseInt(StringUtils.isEmpty(cellStr1) == true ? "0" : cellStr1);
                    // int cell2 = Integer.parseInt(StringUtils.isEmpty(cellStr2) == true ? "0" : cellStr2);
                    String cell = compute(cellStr1, cellStr2, "a+b");
                    //String cell = (cell1+cell2)+"";
                    rowList.add(cell);
                }else if (cid.contains("/")){
                    String[] split = cid.split("/");
                    String cellStr1 = cityMap.get(split[0]);
                    String cellStr2 = cityMap.get(split[1]);
                    String cell = compute(cellStr1, cellStr2, "a/b");
                    if("NULL".equals(cell)){
                        rowList.add(cell);
                    }else {
                        double v = Double.parseDouble(cell);
                        rowList.add(df.format(v));
                    }
                }else if (cid.contains("\\")){
                    String[] split = cid.split("\\\\");
                    String cellStr1 = cityMap.get(split[0]);
                    String cellStr2 = cityMap.get(split[1]);
                    String cell = compute(cellStr1, cellStr2, "a/b");
                    rowList.add(cell);
                }else if (cid.contains("%")){
                    String[] split = cid.split("%");
                    String cell = cityMap.get(split[0]);
                    if(StringUtils.isEmpty(cell)||"NULL".equals(cell)){
                        rowList.add(cell);
                    }else {
                        double v = Double.parseDouble(cell);
                        rowList.add(df.format(v));
                    }
                }else {
                    String cell = cityMap.get(cid);
                    rowList.add(cell);
                }
            }
            tableList.add(rowList);
        }
        return tableList;
    }

    /**
     * 行列转至  外层map 对应的是全球，里面的map是指标统计的key和value
     * @param processBriefs
     * @return
     */
    private Map<String,Map<String, String>> convertCityResulMap(List<PerformanceBriefInfo> processBriefs,String sFlag,String eFalg) {
        Map<String,Map<String, String>> cityMap = new HashMap<>();

        for (String city : citys){
            Map<String, String> cidMap = new HashMap<>();// cid 与word模板对应 C3...
            for (PerformanceBriefInfo pInfo:processBriefs) {
                Integer id = pInfo.getId();
                String valueByCity = BriefHandler.getValueByCity(pInfo, city);
                cidMap.put(sFlag+id+eFalg,valueByCity);
            }
            cityMap.put(city,cidMap);
        }
        return cityMap;
    }

    private String compute(String s1 ,String s2 ,String flag){
        DecimalFormat df;
        Double a;
        Double b;
        try{a = Double.parseDouble(s1);}catch (Exception e){ return "NULL";}
        try{b = Double.parseDouble(s2);}catch (Exception e){ return "NULL";}

        df = new DecimalFormat("0.0000");

        if(flag.equals("a-b")){
            String format = df.format(a - b);
            if (format.indexOf(".0000") != -1){
                return format.substring(0,format.indexOf("."));
            }
            return format;
        }else if(flag.equals("a+b")){
            String format = df.format(a + b);
            if (format.indexOf(".0000") != -1){
                return format.substring(0,format.indexOf("."));
            }
            return format;
        }else if(flag.equals("a/b")){
            if (a==0||b==0) return "0";
            return df.format(a/b);
        }else if(flag.equals("(a-b)/b")){
            if (a-b==0||b==0) return "0";
            return df.format((a-b)/b);
        }else {
            return "NULL";
        }
    }
    private String compute(String s1 ,String s2,String s3 ,String flag){
        Double a;
        Double b;
        Double c;
        try{a = Double.parseDouble(s1);}catch (Exception e){ return "NULL";}
        try{b = Double.parseDouble(s2);}catch (Exception e){ return "NULL";}
        try{c = Double.parseDouble(s3);}catch (Exception e){ return "NULL";}
        DecimalFormat df;
        df = new DecimalFormat("0.0000");
        if(flag.equals("a/b-c")){
            if (a==0||b==0) return df.format(0-c);
            return df.format(a/b-c);
        }else {
            return "NULL";
        }
    }

    /**
     * 根据全球名 返回word列表的行数
     * @param city
     * @return
     */
    public  String getRowNoByCity(String city) {
        for (int i = 0; i <citys.length ; i++) {
            if (citys[i].equals(city)){
                return i+1+"";
            }
        }
        return "0";
    }


    /**
     * 根据每张表的 C编号 返回word列表的列数
     * @param cNo
     * @param columnFortable
     * @return
     */
    public  String getColumnNoByCNo(String cNo,String[] columnFortable) {
        for (int i = 0; i <columnFortable.length ; i++) {
            if (columnFortable[i].equals(cNo)){
                return i+1+"";
            }
        }
        return "0";
    }

    private Map<String,String> table1AddColorData() {
        Map<String,String> data = new HashMap<>();
        /*data.put("3,3","C191C191S(a-b)/b,<0");
        data.put("7,3","C36C36S(a-b)/b,<0");
        data.put("11,3","C68C68S(a-b)/b,<0");
        data.put("3,4","C192C192S(a-b)/b,<0");
        data.put("7,4","C42C42S(a-b)/b,<0");
        data.put("11,4","C74C74S(a-b)/b,<0");
        data.put("3,5","C193C193S(a-b)/b,<0");
        data.put("7,5","C44C44S(a-b)/b,<0");
        data.put("11,5","C76C76S(a-b)/b,<0");*/
        data.put("3,6","C250,>0");
        data.put("6,6","C254,>0");
        data.put("7,6","C53,>0");
        data.put("10,6","C258,>0");
        data.put("11,6","C257,>0");
        data.put("3,7","C261,>0");
        data.put("6,7","C150C260a-b,>0");
        data.put("7,7","C150C150Sa-b,>0");
        data.put("10,7","C153C260a-b,>0");
        data.put("11,7","C153C153Sa-b,>0");
        data.put("3,8","C194C194Sa-b,<0");
        data.put("6,8","C49,<0");
        data.put("7,8","C48,<0");
        data.put("10,8","C81,<0");
        data.put("11,8","C80,<0");
        data.put("3,9","C195C195Sa-b,>0");
        data.put("6,9","C52,>0");
        data.put("7,9","C51,>0");
        data.put("10,9","C84,>0");
        data.put("11,9","C83,>0");
        data.put("3,10","C196C196Sa-b,<0");
        data.put("6,10","C55,<0");
        data.put("7,10","C54,<0");
        data.put("10,10","C87,<0");
        data.put("11,10","C86,<0");
        data.put("3,11","C197C197Sa-b,<0");
        data.put("6,11","C58,<0");
        data.put("7,11","C57,<0");
        data.put("10,11","C90,<0");
        data.put("11,11","C89,<0");
        data.put("3,12","C198C198Sa-b,>0");
        data.put("6,12","C61,>0");
        data.put("7,12","C60,>0");
        data.put("10,12","C93,>0");
        data.put("11,12","C92,>0");
        data.put("3,13","C199C199Sa-b,<0");
        data.put("6,13","C64,<0");
        data.put("7,13","C63,<0");
        data.put("10,13","C96,<0");
        data.put("11,13","C95,<0");
        data.put("3,14","C200C200Sa-b,>0");
        data.put("6,14","C67,>0");
        data.put("7,14","C66,>0");
        data.put("10,14","C99,>0");
        data.put("11,14","C98,>0");
        return data;
    }

    private Map<String,String> table2AddColorData() {
        Map<String,String> data = new HashMap<>();
        data.put("4,6","C267C252a-b,>0");
        data.put("7,6","C271C256a-b,>0");
        data.put("4,7","C205C262C150a/b-c,>0");
        data.put("7,7","C208C263C153a/b-c,>0");
        data.put("4,8","C111C47a-b,<0");
        data.put("7,8","C130C79a-b,<0");
        data.put("4,9","C113C50a-b,>0");
        data.put("7,9","C132C82a-b,>0");
        data.put("4,10","C115C53a-b,<0");
        data.put("7,10","C134C85a-b,<0");
        data.put("4,11","C117C56a-b,<0");
        data.put("7,11","C136C88a-b,<0");
        data.put("4,12","C119C59a-b,>0");
        data.put("7,12","C138C91a-b,>0");
        data.put("4,13","C121C62a-b,<0");
        data.put("7,13","C140C94a-b,<0");
        data.put("4,14","C123C65a-b,>0");
        data.put("7,14","C142C97a-b,>0");
        return data;
    }

    private static boolean alterColor(String value, String flag) {
        Double a;
        try{
            if (value.indexOf("%") != -1){
                a = Double.parseDouble(value.replace("%",""))/100;
            }else{
                a = Double.parseDouble(value);
            }
            if (flag.equals("<0")){
                return a < 0;
            }else if (flag.equals(">0")){
                return a > 0;
            }else if (flag.equals("<0.0001")){
                return a < 0.0001;
            }else if (flag.equals(">0.0001")){
                return a > 0.0001;
            }
            return false;
        }catch (Exception e){
            return false;
        }
    }
}
