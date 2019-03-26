package cn.edu.njupt.outExcel.controller;

import cn.edu.njupt.api.ProjectApplicationControllerApi;
import cn.edu.njupt.utils.ProjectApplicationExportUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.*;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Map;


@RestController
@RequestMapping("/projectApplicationController")
public class ProjectApplicationController implements ProjectApplicationControllerApi {

    @Override
    @PostMapping("/generateProjectApplicationExcel")
    public String generateProjectApplicationExcel(@RequestBody Map map) {
        XSSFWorkbook workBook = null;
        try {
            workBook = ProjectApplicationExportUtil.createWorkBook(map);
            String fileURL="d:/projectexcel/pro"+ System.currentTimeMillis() +".xlsx";
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
