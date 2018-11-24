package com.word.poi.demo.util;

import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileItemFactory;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;


public class ExportToExcelUtils<T> {
    //每次设置导出数量  
    public static int  NUM=1000;  
    public static String title="";  

    /**
     * 解析Excel
     * @param in
     * @param cls
     * @param columns
     * excel表格中字段顺序为:必须要与  columns一一对应
     * @return
     * @throws Exception
     */
    @SuppressWarnings("deprecation")
	public List<T>  excelParseList(InputStream in, Class<T> cls, String[] columns, int sheetNum) throws Exception {
    	List<T> hsTransAccountDtos = new ArrayList<>();
    	//从MultipartFile获取File式的BufferedInputStream
        //BufferedInputStream bf= new BufferedInputStream(multifile.getInputStream());
        XSSFWorkbook workbook= new XSSFWorkbook(in);
        XSSFSheet sheet = workbook.getSheetAt(sheetNum);
        //rowIndex = 1 表头不读取
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {  
            XSSFRow row = sheet.getRow(rowIndex);
            if (row == null) {  
                continue;  
            }  
            XSSFCell cell = null;
            Map<String, String> valMap = new HashMap<>();
            for (short columnIndex = 0; columnIndex < row.getLastCellNum(); columnIndex++) {  
                String value = "";  
                cell = row.getCell(columnIndex);  
                if (cell != null) {  
                    // 注意：如果有乱码，否则可能会出现乱码  
                    // cell.setEncoding(HSSFCell.ENCODING_UTF_16);  
                    switch (cell.getCellType()) {
                    case XSSFCell.CELL_TYPE_STRING:
                        value = cell.getStringCellValue();  
                        break;
                    case XSSFCell.CELL_TYPE_NUMERIC:
                        if (HSSFDateUtil.isCellDateFormatted(cell)) {
                            Date date = cell.getDateCellValue();  
                            if (date != null) {  
                                value = new SimpleDateFormat("yyyy/MM/dd")
                                        .format(date);  
                            } else {  
                                value = "";  
                            }  
                        } else {  
                            value = new DecimalFormat("0").format(cell.getNumericCellValue());  
                        }  
                        break;  
                    case XSSFCell.CELL_TYPE_FORMULA:  
                        // 导入时如果为公式生成的数据则无值  
                        if (!cell.getStringCellValue().equals("")) {  
                            value = cell.getStringCellValue();  
                        } else {  
                            value = cell.getNumericCellValue() + "";  
                        }  
                        break;  
                    case XSSFCell.CELL_TYPE_BLANK:  
                        break;  
                    case XSSFCell.CELL_TYPE_ERROR:  
                        value = "";  
                        break;  
                    case XSSFCell.CELL_TYPE_BOOLEAN:  
                        value = (cell.getBooleanCellValue() == true ? "Y" : "N");  
                        break;  
                    default:  
                        value = "";  
                    }  
                }
                if ( row.getLastCellNum() == 0 && value.trim().equals("")) {
                    break;  
                } 
                valMap.put(columns[columnIndex], value);
            }  
            T  t = setFieldValue(cls, valMap);
            hsTransAccountDtos.add(t);
        }
		return hsTransAccountDtos;
    }
    

    /** 
     * set属性的值到Bean 
     *  
     * @param cls 
     * @param valMap 
     * @throws IllegalAccessException 
     * @throws InstantiationException 
     */  
    private T setFieldValue(Class<T> cls, Map<String, String> valMap) throws InstantiationException, IllegalAccessException {  
        // 取出bean里的所有方法  
        Method[] methods = cls.getDeclaredMethods();  
        Field[] fields = cls.getDeclaredFields();  
        //T bean = (T) cls.getClass(); 
        T bean = cls.newInstance();
        for (Field field : fields) {  
            try {  
                String fieldSetName = parSetName(field.getName());  
                if (!checkSetMet(methods, fieldSetName)) {  
                    continue;  
                }  
                Method fieldSetMet = cls.getMethod(fieldSetName,  
                        field.getType());  
//              String fieldKeyName = parKeyName(field.getName());  
                String  fieldKeyName = field.getName();  
                String value = valMap.get(fieldKeyName);  
                if (null != value && !"".equals(value)) {  
                    String fieldType = field.getType().getSimpleName();  
                    if ("String".equals(fieldType)) {  
                        fieldSetMet.invoke(bean, value);  
                    } else if ("Date".equals(fieldType)) {  
                        Date temp = parseDate(value);  
                        fieldSetMet.invoke(bean, temp);  
                    } else if ("Integer".equals(fieldType)  
                            || "int".equals(fieldType)) {  
                        Integer intval = Integer.parseInt(value);  
                        fieldSetMet.invoke(bean, intval);  
                    } else if ("Long".equalsIgnoreCase(fieldType)) {  
                        Long temp = Long.parseLong(value);  
                        fieldSetMet.invoke(bean, temp);  
                    } else if ("Double".equalsIgnoreCase(fieldType)) {  
                        Double temp = Double.parseDouble(value);  
                        fieldSetMet.invoke(bean, temp);  
                    } else if ("Boolean".equalsIgnoreCase(fieldType)) {  
                        Boolean temp = Boolean.parseBoolean(value);  
                        fieldSetMet.invoke(bean, temp);  
                    } else {  
                        System.out.println("not supper type" + fieldType);  
                    }  
                }  
            } catch (Exception e) {  
                continue;  
            }  
        }
		return bean;  
    }  
    
    
    /** 
     * 拼接在某属性的 set方法 
     *  
     * @param fieldName 
     * @return String 
     */  
    private  String parSetName(String fieldName) {  
        if (null == fieldName || "".equals(fieldName)) {  
            return null;  
        }  
        int startIndex = 0;  
        if (fieldName.charAt(0) == '_')  
            startIndex = 1;  
        return "set"  
                + fieldName.substring(startIndex, startIndex + 1).toUpperCase()  
                + fieldName.substring(startIndex + 1);  
    }  
    
    /** 
     * 格式化string为Date 
     *  
     * @param datestr 
     * @return date 
     */  
    private  Date parseDate(String datestr) {  
        if (null == datestr || "".equals(datestr)) {  
            return null;  
        }  
        try {  
            String fmtstr = null;  
            if (datestr.indexOf(':') > 0) {  
                fmtstr = "yyyy-MM-dd HH:mm:ss";  
            } else {  
                fmtstr = "yyyy-MM-dd";  
            }  
            SimpleDateFormat sdf = new SimpleDateFormat(fmtstr, Locale.UK);  
            return sdf.parse(datestr);  
        } catch (Exception e) {  
            return null;  
        }  
    }  
    
    
    /** 
     * 判断是否存在某属性的 set方法 
     *  
     * @param methods 
     * @param fieldSetMet 
     * @return boolean 
     */  
    private  boolean checkSetMet(Method[] methods, String fieldSetMet) {  
        for (Method met : methods) {  
            if (fieldSetMet.equals(met.getName())) {  
                return true;  
            }  
        }  
        return false;  
    }  
    
    public static FileItem createFileItem(String filePath){
        FileItemFactory factory = new DiskFileItemFactory(16, null);  
        String textFieldName = "textField";  
        int numb = filePath.lastIndexOf(".");  
        String extFile = filePath.substring(numb);  
        FileItem item = factory.createItem(textFieldName, "text/plain", true,  
            "MyFileName" + extFile);  
        File newfile = new File(filePath);  
        int bytesRead = 0;  
        byte[] buffer = new byte[8192];  
        try{  
            FileInputStream fis = new FileInputStream(newfile);  
            OutputStream os = item.getOutputStream();  
            while ((bytesRead = fis.read(buffer, 0, 8192))!= -1){  
                os.write(buffer, 0, bytesRead);  
            }  
            os.close();  
            fis.close();  
        }catch (IOException e){  
            e.printStackTrace();  
        }  
        return item;  
    }

    /**
     * 导出Excel的方法
     * @param headers 表头
     * @param columns 列名
     * @param result 结果集
     * @param out 输出流
     * @param fileName 文件名
     * @param request
     * @param response
     * @throws Exception
     */
    public void exportExcel(String[] headers, String[] columns, List<T> result, OutputStream out, String fileName, HttpServletRequest request, HttpServletResponse response) throws Exception{

        setResponseHeader(response,fileName);
        // 声明一个工作薄
        SXSSFWorkbook workbook = new SXSSFWorkbook(1000);
        // 生成一个表格
        Sheet sheet = workbook.createSheet(title);
        // 设置表格默认列宽度为18个字节
        sheet.setDefaultColumnWidth((short)18);

        // 生成一个样式
        CellStyle style = workbook.createCellStyle();
        // 设置这些样式
        style.setFillForegroundColor(HSSFColor.GOLD.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);

        // 生成一个字体
        Font font = workbook.createFont();
        font.setColor(HSSFColor.VIOLET.index);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        // 把字体应用到当前的样式
        style.setFont(font);
        // 指定当单元格内容显示不下时自动换行
        style.setWrapText(true);

        Row row = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = row.createCell((short)i);
            cell.setCellStyle(style);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        // 遍历集合数据，产生数据行
        if(result != null){
            int index = 1;
            for(T t:result){
                row = sheet.createRow(index);
                index++;
                for(short i = 0; i < columns.length; i++) {
                    Cell cell = row.createCell(i);
                    String fieldName = columns[i];
                    String getMethodName = "get"
                            + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
                    Class tCls = t.getClass();
                    Method getMethod = tCls.getMethod(getMethodName, new Class[]{});
                    Object value = getMethod.invoke(t, new Class[]{});
                    String textValue = null;
                    if(value == null) {
                        textValue = "";
                        XSSFRichTextString richString = new XSSFRichTextString(textValue);
                        cell.setCellValue(richString);
                    }else if (value instanceof Date) {
                        Date date = (Date) value;
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy/mm/dd");
                        textValue = sdf.format(date);
                        XSSFRichTextString richString = new XSSFRichTextString(textValue);
                        cell.setCellValue(richString);
                    }else if (value instanceof BigDecimal) {
                        cell.setCellValue(Double.parseDouble(value.toString()));
                    }else if (value instanceof Long) {
                        cell.setCellValue(Double.parseDouble(value.toString()));
                    }else{
                        //其它数据类型都当作字符串简单处理
                        textValue = value.toString();
                        XSSFRichTextString richString = new XSSFRichTextString(textValue);
                        cell.setCellValue(richString);
                    }
                }
            }
        }
        workbook.write(out);
    }

    /** 设置响应头 */
    private void setResponseHeader(HttpServletResponse response, String fileName) {
        try {
            this.title=fileName;
            response.reset();// 清空输出流
            response.setContentType("application/octet-stream;charset=UTF-8");
            response.setHeader("Content-Disposition", "attachment;filename="
                    +new String(this.title.getBytes("GB2312"), "8859_1")
                    + ".xlsx");
            response.addHeader("Pargam", "no-cache");
            response.addHeader("Cache-Control", "no-cache");
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
}