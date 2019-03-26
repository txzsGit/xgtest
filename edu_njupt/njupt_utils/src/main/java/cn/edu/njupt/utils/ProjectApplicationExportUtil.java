package cn.edu.njupt.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 项目申请表工具类
 */
public class ProjectApplicationExportUtil {
    //项目申请表模板
    private static final String templateFilePath =
            "C:\\template\\projectApplicationTemplate.xlsx";


    //将模板表格导入程序
    public static XSSFWorkbook createWorkBook(Map<String, String> map)
            throws Exception
    {
        InputStream in = new FileInputStream(new File(templateFilePath));
        XSSFWorkbook work = new XSSFWorkbook(in);
        // 得到excel的第0张表
        XSSFSheet sheet = work.getSheetAt(0);


        replaceExcelData(sheet, map,work);

        return work;
    }

    public static void replaceExcelData(XSSFSheet sheet, Map<String, String> map ,XSSFWorkbook workbook)
    {
        int rowNum = sheet.getLastRowNum();
        for (int i = 0; i <= rowNum; i++)
        {//行
            XSSFRow row = sheet.getRow(i);
            if(i==1){//表头单独设置
                XSSFCell cell =row.getCell(0);
                cell.setCellValue(map.get("head"));//设置单元格内容
                XSSFCellStyle cellStyle = workbook.createCellStyle();//设置样式
                cellStyle.setAlignment(HorizontalAlignment.CENTER);//横向居中
                cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//纵向居中
                //生成一个字体
                XSSFFont font = workbook.createFont();
                font.setFontHeightInPoints((short) 16);
                font.setFontName("黑体");
                cellStyle.setFont(font);
                cell.setCellStyle(cellStyle);
            }
            if (row == null)
                continue;
            for (int j = 0; j < row.getPhysicalNumberOfCells(); j++)
            {//列
                XSSFCell cell = row.getCell(j);
                if (cell == null)
                    continue;
                String key = cell.getStringCellValue();
                //System.out.println(key);
                if (map.containsKey(key)){
                    if(key!="head"){
                        cell.setCellValue(map.get(key));
                    }
                }
            }
        }
    }

    public static void main(String[] args) throws Exception {
        Map map = new HashMap();
        map.put("head","二零一九年“八小、三小”活动实施项目推荐评比表");
        map.put("company","徐工集团");
        map.put("date","2019.03.23");
        map.put("projectName","徐工调研");
        map.put("projectNum","001");
        map.put("header","杨总");
        map.put("teamers","小王、小李");
        map.put("projectType","科研");
        map.put("projectTime","2018.02.12");
        map.put("beginTime","2018.03.01");
        map.put("projectTime","2018.02.12");
        map.put("finishTime","2018.05.01");
        map.put("projectReson","立项缘由：\n" +
                "不详");
        map.put("measure","项目衡量指标：\n" +
                "不详");
        map.put("nowLevel","现有指标水平：\n" +
                "不详");
        map.put("targetLevel","目标指标水平：\n" +
                "不详");
        map.put("method","实施方法：\n" +
                "待定");
        map.put("effect","实施效果：                                                                                          良好");
        map.put("降本节能   save       质量提升     ascension   提高工效  effici倍     经济效益        economic","降本节能   √         质量提升     √        提高工效  10倍          经济效益       10000");
        map.put("advice","同意");
        map.put("adviceDate","2018.2.20");
        System.out.println(map);
        XSSFWorkbook workBook = createWorkBook(map);
        OutputStream stream = new FileOutputStream("d:/2.xlsx");
        workBook.write(stream);
        workBook.close();
    }
}
