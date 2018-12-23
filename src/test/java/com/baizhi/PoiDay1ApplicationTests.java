package com.baizhi;

import com.baizhi.entity.User;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

@RunWith(SpringRunner.class)
@SpringBootTest
public class PoiDay1ApplicationTests {

    @Test
    public void contextLoads() {


    }
    @Test
    public void demo1() throws IOException {
        //1.创建一个对象 创建一个Excal文件对象  HSSFWorkbook
        HSSFWorkbook workbook = new HSSFWorkbook();
        //2.新建一张工作表 HSSFSheet
        HSSFSheet user = workbook.createSheet("user");
        //3.写入一行数据
        //(1)创建一行 HSSRow  参数:创建第几行 从0开始
        HSSFRow row = user.createRow(0);
        //(2)创建一个单元格
        HSSFCell cell = row.createCell(0);
        //(3)在单元格中写入数据
        cell.setCellValue("测试数据");
        //4.输出
        workbook.write(new FileOutputStream(new File("F://aa.xls")));
    }

    //第二个  设置标题栏
    @Test
    public void demo2() throws IOException {
        List<String> title = Arrays.asList("编号","用户名","密码");
        //1.创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //2.创建一个工作表
        HSSFSheet user = workbook.createSheet("user");
        //3.写入数据
        //(1)创建一行
        HSSFRow row = user.createRow(0);
        //(2)创建3个单元格 分别写入数据
        for (int i = 0; i <title.size() ; i++) {
            HSSFCell cell = row.createCell(i);
            //获取集合中的元素
            cell.setCellValue(title.get(i));
        }
        //4.输出到本地
        workbook.write(new FileOutputStream(new File("F://b.xls")));
    }


}
