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
public class WeekPlanExportUtil {
    //项目申请表模板
    private static final String templateFilePath =
            "C:\\template\\weekPlanTemplate.xlsx";


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
                font.setFontHeightInPoints((short) 14);
                font.setFontName("黑体");
                cellStyle.setFont(font);
                cell.setCellStyle(cellStyle);
            }
            if(i==2){//日期单独设置
                for(int t=1;t<=7;t++){
                    XSSFCell cell =row.getCell(t+2);
                    cell.setCellValue((String)map.get(t+""));//设置单元格内容
                    XSSFCellStyle cellStyle = workbook.createCellStyle();//设置样式
                    cellStyle.setAlignment(HorizontalAlignment.CENTER);//横向居中
                    cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//纵向居中
                    //生成一个字体
                    XSSFFont font = workbook.createFont();
                    font.setFontHeightInPoints((short) 9);
                    font.setFontName("仿宋_GB2312");
                    cellStyle.setFont(font);
                    cell.setCellStyle(cellStyle);
                }
            }

            List list =(List<Map>) map.get("list");
            if(i==rowNum){//最后一行
                if(list!=null&&list.size()>1) {
                    int size = list.size();
                    for (int m = 1; m < size; m++) {
                        sheet.copyRows(4, 4, 4 + m, new CellCopyPolicy());
                    }
                    fillData(sheet,size,list);
                }
            }



        }
    }

    public static void fillData(XSSFSheet sheet,int size,List list){
        for(int i=0;i<size;i++) {
            XSSFRow row = sheet.getRow(4+i);
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
        map.put("head","2019年3月18第三周执行计划");
        map.put("1","3-18");
        map.put("2","3-19");
        map.put("3","3-19");
        map.put("4","3-20");
        map.put("5","3-21");
        map.put("6","3-22");
        map.put("7","3-23");
        ArrayList<Object> list = new ArrayList<>();
        Map map1=new HashMap();
        map1.put("num","1");
        map1.put("project","徐工");
        map1.put("content","徐工");
        map1.put("dept","人事部");
        map1.put("collab","111");
        map1.put("result","222");
        map1.put("supe","333");
        map1.put("guide","444");
        map1.put("instr","555");
        map1.put("content1","qqq");
        map1.put("content2","www");
        map1.put("content3","eee");
        map1.put("content4","rrr");
        map1.put("content5","ttt");
        map1.put("content6","yyy");
        map1.put("content7","uuu");
        Map map2=new HashMap();
        map2.put("num","2");
        map2.put("project","徐工");
        map2.put("content","徐工");
        map2.put("dept","人事部");
        map2.put("collab","111");
        map2.put("result","222");
        map2.put("supe","333");
        map2.put("guide","444");
        map2.put("instr","555");
        map2.put("content1","qqq");
        map2.put("content2","www");
        map2.put("content3","eee");
        map2.put("content4","rrr");
        map2.put("content5","ttt");
        map2.put("content6","yyy");
        map2.put("content7","uuu");
        Map map3=new HashMap();
        map3.put("num","3");
        map3.put("project","徐工");
        map3.put("content","徐工");
        map3.put("dept","人事部");
        map3.put("collab","111");
        map3.put("result","222");
        map3.put("supe","333");
        map3.put("guide","444");
        map3.put("instr","555");
        map3.put("content1","qqq");
        map3.put("content2","www");
        map3.put("content3","eee");
        map3.put("content4","rrr");
        map3.put("content5","ttt");
        map3.put("content6","yyy");
        map3.put("content7","uuu");
        list.add(map1);
        list.add(map2);
        list.add(map3);
        map.put("list",list);
        XSSFWorkbook workBook = createWorkBook(map);
        OutputStream stream = new FileOutputStream("C:\\Users\\txzs\\Desktop\\6.xlsx");
        workBook.write(stream);
        workBook.close();
    }
}
