package cn.edu.njupt.api;

import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;

import java.util.Map;

/**
 * 周计划表
 */
@Api(value="周计划表接口",description = "生成周计划表")
public interface WeekPlanControllerApi {
    //生成申请表excel表格文件
    @ApiOperation("根据map生成周计划表格")
    public  String  generateWeekPlanExcel(Map map);
}
