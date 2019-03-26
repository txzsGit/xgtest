package cn.edu.njupt.outExcel.controller;

import cn.edu.njupt.api.MonthPlanControllerApi;
import cn.edu.njupt.api.WeekPlanControllerApi;
import cn.edu.njupt.utils.MonthPlanExportUtil;
import cn.edu.njupt.utils.WeekPlanExportUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Map;
import java.util.UUID;


@RestController
@RequestMapping("/weekPlanController")
public class WeekPlanController implements WeekPlanControllerApi {
    @Override
    @GetMapping("/generateWeekPlanExcel")
    public String generateWeekPlanExcel(@RequestBody Map map) {
        XSSFWorkbook workBook = null;
        try {
            workBook = WeekPlanExportUtil.createWorkBook(map);
            String fileURL="d:/weekexcel/week"+ System.currentTimeMillis()+".xlsx";
            OutputStream stream = new FileOutputStream(fileURL);
            workBook.write(stream);
            workBook.close();
            return fileURL;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return  null;
    }
}
