package cn.edu.njupt.api;

import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;

import java.util.Map;

/**
 * 项目申请表api
 */
@Api(value="项目申请表管理接口",description = "生成项目申请表")
public interface ProjectApplicationControllerApi {
    //生成申请表excel表格文件
    @ApiOperation("根据map生成项目申请表格")
    public  String  generateProjectApplicationExcel(Map map);
}
