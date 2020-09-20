package com.im;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.im.modules.dao.UserDao;
import com.im.modules.easy.DemoDataListener;
import com.im.modules.easy.DemoDataListener2;
import com.im.modules.entities.DemoData;
import com.im.modules.entities.User;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.List;

/**
 *  文档参照 ： https://www.yuque.com/easyexcel/doc/easyexcel
 */
@SpringBootTest
public class EasyExcelTests {

    // 存放的路径
    String path = "F:\\one\\";

    @Autowired
    private UserDao userDao;

    // EasyExcel 导入 Excel
    // 1. 实体类 给每个字段加@注解 @ExcelProperty
    @Test
    void testWrite() throws Exception{

        // List 实体类数据
        List<User> user = userDao.getUser();
        System.out.println(user);

        // 写法1  写入 Excel
        String fileName = path + "EasyExcel.xlsx";
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        // 如果这里想使用03 则 传入excelType参数即可
        EasyExcel.write(fileName, User.class).sheet("标题1").doWrite(user);

        // 写法2
//        String fileName = path + "EasyExcel.xlsx";
//        // 这里 需要指定写用哪个class去写
//        ExcelWriter excelWriter = null;
//        try {
//            excelWriter = EasyExcel.write(fileName, User.class).build();
//            WriteSheet writeSheet = EasyExcel.writerSheet("模板").build();
//            excelWriter.write(user, writeSheet);
//        } finally {
//            // 千万别忘记finish 会帮忙关闭流
//            if (excelWriter != null) {
//                excelWriter.finish();
//            }
//        }

    }


    // EasyExcel 写入 Excel
    // 需要 DemoDataListener监听器
    //  Dao 层
    @Test
    void testRead() throws Exception{

        String fileName = path + "EasyExcel.xlsx";
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();

    }

    @Test
    void testRead2() throws Exception{

        String fileName = path + "EasyExcel3.xlsx";
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        EasyExcel.read(fileName, User.class, new DemoDataListener2()).sheet().doRead();

    }

}
