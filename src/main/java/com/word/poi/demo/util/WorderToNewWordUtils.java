package com.word.poi.demo.util;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 通过word模板生成新的word工具类
 * 
 */
public class WorderToNewWordUtils {

    /**
     * 根据模板生成新word文档
     * 判断表格是需要替换还是需要插入，判断逻辑有${}为替换，表格无${}为插入
     * @param textMap 需要替换的信息集合
     * @return 成功返回true,失败返回false
     */
    public static XWPFDocument changWordForcommon(InputStream in,Map<String, String> textMap) {
        XWPFDocument document = null;
        try {
            //第一步：获取docx解析对象
            document = new XWPFDocument(in);
            //第二步：重构XWPFDocument中 不合格的 占位符 比如：${abcd}分开为${a，b..,cd}
            refactorXWPFDocument(document);
            //第三步：解析替换文本段落对象
            changeText(document, textMap);
            //第四步：解析替换表格对象
           changeTableCity(document, textMap);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return document;

    }
    /**
     * 根据模板生成新word文档
     * 判断表格是需要替换还是需要插入，判断逻辑有${}为替换，表格无${}为插入
     * @param textMap 需要替换的信息集合
     * @param extendTableMap 需要插入的表格信息集合
     * @param colorMapForCell
     * @return 成功返回true,失败返回false
     */
    public static XWPFDocument changWordForComplex(
            InputStream in,
            Map<String, String> textMap, 
            Map<String, List<List<String>>> extendTableMap,
            Map<String, List<String>> colorMapForCell) {
        XWPFDocument document = null;
        try {
            //第一步：获取docx解析对象
            document = new XWPFDocument(in);
            //第二步：重构XWPFDocument中 不合格的 占位符 比如：${abcd}分开为${a，b..,cd}
            refactorXWPFDocument(document);
            //第三步：解析替换文本段落对象
           changeText(document, textMap);
            //第四步：解析替换表格对象
           changeTableProvince(document, textMap, extendTableMap);
            //第五步：根据逻辑设置表格cell的背景颜色
            setColorForCell(document,colorMapForCell);

        } catch (IOException e) {
            e.printStackTrace();
        }
        return document;

    }
    private static void setColorForCell(XWPFDocument document, Map<String, List<String>> colorMapForCell) {
        List<XWPFTable> tables = document.getTables();
        for (int i = 0; i <tables.size() ; i++) {
            List<String> placeList = colorMapForCell.get("table" + i);//获取表格坐标list
            if (CollectionUtils.isNotEmpty(placeList)){
                for (String palce:placeList) {
                    String[] split = palce.split(",");
                    int x = Integer.parseInt(split[0]);
                    int y = Integer.parseInt(split[1]);
                    tables.get(i).getRow(x).getCell(y).setColor("FF0000");
                }
            }
        }
    }

    /**
     * 替换段落文本
     * @param document docx解析对象
     * @param textMap 需要替换的信息集合
     */
    public static void changeText(XWPFDocument document, Map<String, String> textMap){
        //获取段落集合
        List<XWPFParagraph> paragraphs = document.getParagraphs();

        for (XWPFParagraph paragraph : paragraphs) {
            //判断此段落时候需要进行替换
            String text = paragraph.getText();
            //重构段落  不合格的 占位符 比如：${abcd}分开为${a，b..,cd}
           // refactorParagraph(paragraph);
            if(checkText(text)){
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    //替换模板原来位置
                    run.setText(changeValue(run.toString(), textMap),0);
                }
            }
        }

    }

    /**
     * 替换表格对象方法
     * @param document docx解析对象
     * @param textMap 需要替换的信息集合
     */
    public static void changeTableCity(XWPFDocument document, Map<String, String> textMap
           ){
        //获取表格对象集合
        List<XWPFTable> tables = document.getTables();
        for (int i = 0; i < tables.size(); i++) {
            //只处理行数大于等于2的表格，且不循环表头
            XWPFTable table = tables.get(i);
            if(table.getRows().size()>1){
                //判断表格是需要替换还是需要插入，判断逻辑有$为替换，表格无$为插入
                if(checkText(table.getText())){
                    List<XWPFTableRow> rows = table.getRows();
                    //遍历表格,并替换模板
                    eachTable(rows, textMap);
                }
            }
        }
    }
    /**
     * 替换表格对象方法
     * @param document docx解析对象
     * @param textMap 需要替换的信息集合
     * @param extendTableMap 需要插入的表格信息集合
     */
    public static void changeTableProvince(XWPFDocument document, Map<String, String> textMap,
                                           Map<String, List<List<String>>> extendTableMap){
        //获取表格对象集合
        List<XWPFTable> tables = document.getTables();
        int tableNo = 0;
        for (int i = 0; i < tables.size(); i++) {
            //只处理行数大于等于2的表格，且不循环表头
            XWPFTable table = tables.get(i);
            if(table.getRows().size()>1){
                //判断表格是需要替换还是需要插入，判断逻辑有$为替换，表格无$为插入
                if(checkText(table.getText())){
                    List<XWPFTableRow> rows = table.getRows();
                    //遍历表格,并替换模板
                    eachTable(rows, textMap);
                    tableNo++;
                }else{
                    List<List<String>> tableList = extendTableMap.get("table" + tableNo);
                    insertTable(table, tableList);
                    tableNo++;
                }
            }
        }
    }





    /**
     * 遍历表格
     * @param rows 表格行对象
     * @param textMap 需要替换的信息集合
     */
    public static void eachTable(List<XWPFTableRow> rows ,Map<String, String> textMap){
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                //判断单元格是否需要替换
                if(checkText(cell.getText())){
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    for (XWPFParagraph paragraph : paragraphs) {
                        List<XWPFRun> runs = paragraph.getRuns();
                        for (XWPFRun run : runs) {
                            run.setText(changeValue(run.toString(), textMap),0);
                        }
                    }
                }
            }
        }
    }

    /**
     * 为表格插入数据，行数不够添加新行
     * @param table 需要插入数据的表格
     * @param tableList 插入数据集合
     */
    public static void insertTable(XWPFTable table, List<List<String>> tableList){
        //创建行,根据需要插入的数据添加新行，不处理表头
        for(int i = 1; i < tableList.size(); i++){
            //XWPFTableRow row =table.createRow();
        }
        //遍历表格插入数据
        List<XWPFTableRow> rows = table.getRows();
        for(int i = 1; i < rows.size(); i++){
            XWPFTableRow newRow = table.getRow(i);
            List<XWPFTableCell> cells = newRow.getTableCells();
            for(int j = 0; j < cells.size(); j++){
                XWPFTableCell cell = cells.get(j);
               // cell.setText(tableList.get(i-1).get(j));
                XWPFParagraph xwpfParagraph = cell.getParagraphs().get(0);//获取cell中的第一个段落第一个ruan并设置字体样式和内容
                XWPFRun run = xwpfParagraph.createRun();
                run.setFontSize(9);
                run.setFontFamily("宋体");
                run.setText(tableList.get(i-1).get(j));
            }
        }
    }



    /**
     * 判断文本中时候包含{{
     * @param text 文本
     * @return 包含返回true,不包含返回false
     */
    public static boolean checkText(String text){
        boolean check  =  false;
        if(null != text && text.indexOf("$")!= -1){
            check = true;
        }
        return check;

    }

    /**
     * 匹配传入信息集合与模板
     * @param value 模板需要替换的区域
     * @param textMap 传入信息集合
     * @return 模板需要替换区域信息集合对应值
     */
    public static String changeValue(String value, Map<String, String> textMap){
        Set<Entry<String, String>> textSets = textMap.entrySet();
        for (Entry<String, String> textSet : textSets) {
            //匹配模板与替换值 格式{{key}}
            String key = "${"+textSet.getKey()+"}";

            if(null != value && value.indexOf(key)!= -1){
                String s = textSet.getValue();
                value = s;
            }
        }
        //模板未匹配到区域替换为空
        if(checkText(value)){
            value = "NULL";
        }
        return value;
    }

    public static void refactorXWPFDocument(XWPFDocument doc) {
        try {
            buildParagraph(doc);
            buildTable(doc);
            /*List<XWPFParagraph> paragraphs = doc.getParagraphs();
            for (XWPFParagraph paragraph : paragraphs){
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs){
                    System.out.println(run.toString());
                }
            }*/
        } catch (Exception e) {
            e.printStackTrace();
        }
        //return doc;
    }

    public static void refactorParagraph(XWPFParagraph paragraph) {
        int start = -1;
        int end = -1;
        List<XWPFRun> runs = paragraph.getRuns();
        for (int i = 0; i < runs.size() ; i++) {
            String runText = runs.get(i).toString();
            if ('$' == runText.charAt(0)&&'}' == runText.charAt(runText.length() - 1)){
                continue;
            }
            if ('$' == runText.charAt(0)){
                start = i;
            }
            if ('}' == runText.charAt(runText.length() - 1)){
                end = i;
                break;
            }
        }
        if (start != -1){
            mergeRun(paragraph,start,end);
            refactorParagraph(paragraph);
        }
    }

  /*  public static void main(String[] args) {
        String runText ="";
        char c = runText.charAt(0);
        System.out.println(c);

    }*/
    public static void mergeRun(XWPFParagraph paragraph, int start, int end) {
        int removeCount = end-start;//删除次数
        int removeIndex = start+1;//删除开始位置

        List<XWPFRun> runs = paragraph.getRuns();
        for (int i = 0; i <runs.size() ; i++) {
            System.out.println(runs.get(i).toString());
        }
        StringBuilder sb = new StringBuilder();
        sb.append(runs.get(start).toString());

        for (int i = 0; i < removeCount; i++){
            sb.append(runs.get(removeIndex).toString());
            paragraph.removeRun(removeIndex);
        }
        runs.get(start).setText(sb.toString(),0);
    }

    public static Matcher matcher(String str) {
        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }


    public static void buildParagraph(XWPFDocument document){
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            if (matcher(paragraph.getText()).find()) {
                removeEmptyRun(paragraph);
                refactorParagraph(paragraph);
            }
        }
    }
    public static void buildTable(XWPFDocument document){
        List<XWPFTable> tables = document.getTables();
        for (XWPFTable xwpfTable : tables) {
            List<XWPFTableRow> rows = xwpfTable.getRows();
            for (XWPFTableRow row :rows){
                List<XWPFTableCell> tableCells = row.getTableCells();
                for (XWPFTableCell tableCell : tableCells){
                    List<XWPFParagraph> paragraphs = tableCell.getParagraphs();
                    for (XWPFParagraph paragraph : paragraphs){
                        if (matcher(paragraph.getText()).find()) {
                            removeEmptyRun(paragraph);
                            refactorParagraph(paragraph);
                        }
                    }
                }
            }

        }
    }
    private static void removeEmptyRun(XWPFParagraph paragraph) {
        List<XWPFRun> runsq = paragraph.getRuns();
        for (int i = 0; i <runsq.size() ; i++) {
            String runText = runsq.get(i).toString();
            if (StringUtils.isEmpty(runText)){
                paragraph.removeRun(i);
                break;
            }
        }
        for (int i = 0; i <runsq.size() ; i++) {
            String runText = runsq.get(i).toString();
            if (StringUtils.isEmpty(runText)){
                removeEmptyRun(paragraph);
                break;
            }
        }
    }

}

