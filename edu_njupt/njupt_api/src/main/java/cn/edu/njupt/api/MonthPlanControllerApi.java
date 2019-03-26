package cn.edu.njupt.api;

import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;

import java.util.Map;

/**
 * 月季度表
 */
@Api(value="月季度表接口",description = "生成月季度计划表")
public interface MonthPlanControllerApi {
    //生成申请表excel表格文件
    @ApiOperation("根据map生成月季度计划表格")
    public  String  generateMonthPlanExcel(Map map);
}
