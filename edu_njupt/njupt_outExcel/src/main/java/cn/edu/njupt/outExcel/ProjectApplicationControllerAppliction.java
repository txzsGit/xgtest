package cn.edu.njupt.outExcel;

/**
 * @author ：TengXun
 * @date ：Created in 2019/3/22 16:49
 */

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.ComponentScan;

/**
 * springboot启动类
 */
@SpringBootApplication
@ComponentScan(basePackages={"cn.edu.njupt.api"})//扫描接口
@ComponentScan(basePackages={"cn.edu.njupt.outExcel"})//扫描本项目下的所有类
public class ProjectApplicationControllerAppliction {
    public static void main(String[] args) {
        SpringApplication.run(ProjectApplicationControllerAppliction.class);
    }
}
