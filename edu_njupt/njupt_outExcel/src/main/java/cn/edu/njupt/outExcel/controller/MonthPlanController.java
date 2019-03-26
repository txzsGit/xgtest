package cn.edu.njupt.outExcel.controller;

import cn.edu.njupt.api.MonthPlanControllerApi;
import cn.edu.njupt.utils.MonthPlanExportUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.*;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Map;



@RestController
@RequestMapping("/monthPlanController")
public class MonthPlanController implements MonthPlanControllerApi {
    @Override
    @PostMapping("/generateMonthPlanExcel")
    public String generateMonthPlanExcel(@RequestBody Map map) {
                XSSFWorkbook workBook = null;
                try {
                    workBook = MonthPlanExportUtil.createWorkBook(map);
                    String fileURL="d:/monthexcel/mon"+ System.currentTimeMillis() +".xlsx";
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
