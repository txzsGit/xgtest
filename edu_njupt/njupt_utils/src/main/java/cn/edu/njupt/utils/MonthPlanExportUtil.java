package cn.edu.njupt.utils;

import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 项目申请表工具类
 */
public class MonthPlanExportUtil {
    //项目申请表模板
    private static final String templateFilePath =
            "C:\\template\\monthPlanTemplate.xlsx";


    //将模板表格导入程序
    public static XSSFWorkbook createWorkBook(Map map)
            throws Exception
    {
        InputStream in = new FileInputStream(new File(templateFilePath));
        XSSFWorkbook work = new XSSFWorkbook(in);
        // 得到excel的第0张表
        XSSFSheet sheet = work.getSheetAt(0);


        replaceExcelData(sheet,map,work);

        return work;
    }

    public static void replaceExcelData(XSSFSheet sheet, Map<String,Object> map,XSSFWorkbook workbook)
    {
        int rowNum = sheet.getLastRowNum();
        for (int i = 0; i <= rowNum; i++)
        {
            XSSFRow row = sheet.getRow(i);
            if (row == null)
                continue;
            if(i==0){//表头单独设置
                XSSFCell cell =row.getCell(0);
                cell.setCellValue((String)map.get("head"));//设置单元格内容
                XSSFCellStyle cellStyle = workbook.createCellStyle();//设置样式
                cellStyle.setAlignment(HorizontalAlignment.CENTER);//横向居中
                cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//纵向居中
                //生成一个字体
                XSSFFont font = workbook.createFont();
                font.setFontHeightInPoints((short) 18);
                font.setFontName("黑体");
                cellStyle.setFont(font);
                cell.setCellStyle(cellStyle);
            }
            List list =(List<Map>) map.get("list");
            if(i==rowNum){//最后一行
                if(list!=null&&list.size()>1) {
                    int size = list.size();
                    for (int m = 1; m < size; m++) {
                        sheet.copyRows(2, 2, 2 + m, new CellCopyPolicy());
                    }
                    fillData(sheet,size,list);
                }
            }



        }
    }

    public static void fillData(XSSFSheet sheet,int size,List list){
        for(int i=0;i<size;i++) {
            XSSFRow row = sheet.getRow(2+i);
            for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {//列
                XSSFCell cell = row.getCell(j);
                if (cell == null)
                    continue;
                String key = cell.getStringCellValue();
                // System.out.println(key);
                    Map map1 = (Map) list.get(i);
                    if (map1.containsKey(key)) {
                        cell.setCellValue((String) map1.get(key));
                    }
                }
            }
        }
    public static void main(String[] args) throws Exception {
        Map map=new HashMap();
        map.put("head","2019年3月党务工作执行计划");
        ArrayList<Object> list = new ArrayList<>();
        Map map1=new HashMap();
        map1.put("num","1");
        map1.put("project","徐工");
        map1.put("content","徐工");
        map1.put("action","徐工");
        map1.put("target","徐工");
        map1.put("startTime","2019.1.1");
        map1.put("finishTime","2019.2.2");
        map1.put("dept","人事部");
        map1.put("reps","111");
        Map map2=new HashMap();
        map2.put("num","2");
        map2.put("project","徐工");
        map2.put("content","徐工");
        map2.put("action","徐工");
        map2.put("target","徐工");
        map2.put("startTime","2019.1.1");
        map2.put("finishTime","2019.2.2");
        map2.put("dept","人事部");
        map2.put("reps","111");
        Map map3=new HashMap();
        map3.put("num","3");
        map3.put("project","徐工");
        map3.put("content","徐工");
        map3.put("action","徐工");
        map3.put("target","徐工");
        map3.put("startTime","2019.1.1");
        map3.put("finishTime","2019.2.2");
        map3.put("dept","人事部");
        map3.put("reps","111");
        list.add(map1);
        list.add(map2);
        list.add(map3);
        map.put("list",list);
        System.out.println(map);
        XSSFWorkbook workBook = createWorkBook(map);
        OutputStream stream = new FileOutputStream("C:\\Users\\txzs\\Desktop\\6.xlsx");
        workBook.write(stream);
        workBook.close();
    }
}
