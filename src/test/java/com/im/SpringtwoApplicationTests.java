package com.im;

import com.im.modules.dao.UserDao;
import com.im.modules.entities.User;
import lombok.experimental.Accessors;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.*;


@SpringBootTest
public class SpringtwoApplicationTests {

    // 存放的路径
    String path = "F:\\one\\";

    @Autowired
    private UserDao userDao;

    // 2003 版本和 2007 版本存在兼容性的问题！03最多只有 65535 行
    // 导出 03版本的 Excel  HSSF
    @Test
    void testWrite03() throws Exception{

        // 创建新的Excel 工作簿
        Workbook workbook = new HSSFWorkbook();
        // 创建工作表
        Sheet sheet = workbook.createSheet("03统计表");
        // 创建行（row 1）
        Row row1 = sheet.createRow(0);
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日浏览量");
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(999);
        // 创建行（row 2）
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        Cell cell22 = row2.createCell(1);
        String dateTime = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(dateTime);
        // 新建一输出文件流（注意：要先创建文件夹）
        FileOutputStream out = new FileOutputStream(path + "观众统计表 03.xls");
        // 把相应的Excel 工作簿存盘
        workbook.write(out);
        // 操作结束，关闭文件
        out.close();
        System.out.println("文件生成成功");

    }
    // 导出 07版本 Excel    XSSF
    @Test
    void testWrite07() throws Exception{

        // 数据
        List<User> user = userDao.getUser();
        System.out.println(user);

        // 创建新的Excel 工作簿
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("07统计表");

        Row row1 = sheet.createRow(0);
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("ID值");
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("用户名");
        Cell cell13 = row1.createCell(2);
        cell13.setCellValue("真实姓名");
        Cell cell14 = row1.createCell(3);
        cell14.setCellValue("性别");

        for (int i = 0; i < user.size(); i++) {
            Row row = sheet.createRow(i+1);
            User u = user.get(i);
            Cell cell = row.createCell(0);
            cell.setCellValue(u.getId());
            Cell cell1 = row.createCell(1);
            cell1.setCellValue(u.getUsername());
            Cell cell2 = row.createCell(2);
            cell2.setCellValue(u.getRealName());
            Cell cell3 = row.createCell(3);
            cell3.setCellValue(u.getGender());
        }

        // 新建一输出文件流（注意：要先创建文件夹）
        FileOutputStream out = new FileOutputStream(path + "用户统计表 07.xlsx");
        // 把相应的Excel 工作簿存盘
        workbook.write(out);
        // 操作结束，关闭文件
        out.close();
        System.out.println("文件生成成功");
    }

    // 导出 07版本 大文件  HSSF，需要时间长
    @Test
    void testWrite07BigData() throws Exception{

        //记录开始时间
        long begin = System.currentTimeMillis();
        //创建一个XSSFWorkbook
        Workbook workbook = new XSSFWorkbook();
        //创建一个sheet
        Sheet sheet = workbook.createSheet("表标题");
        //xls文件最大支持65536行
        for (int rowNum = 0; rowNum < 100000; rowNum++) {
            //创建一个行
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                //创建单元格
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("结束...");
        FileOutputStream out = new FileOutputStream(path + "bigdata07.xlsx");
        workbook.write(out);
        // 操作结束，关闭文件
        out.close();
        //记录结束时间
        long end = System.currentTimeMillis();
        System.out.println((double)(end - begin)/1000);

    }

    // 导出 07版本 大文件  SXSSF，需要时间短
    // 过程中会产生临时文件，需要清理临时文件
    // 如果想自定义内存中数据的数量，可以使用new SXSSFWorkbook ( 数量 )
    @Test
    void testWrite07BigDataFast() throws Exception{

        //记录开始时间
        long begin = System.currentTimeMillis();
        //创建一个SXSSFWorkbook
        Workbook workbook = new SXSSFWorkbook();
        //创建一个sheet
        Sheet sheet = workbook.createSheet();
        //xls文件最大支持65536行
        for (int rowNum = 0; rowNum < 100000; rowNum++) {
        //创建一个行
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {//创建单元格
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("done");
        FileOutputStream out = new FileOutputStream(path + "bigdata07-fast.xlsx");
        workbook.write(out);
        // 操作结束，关闭文件
        out.close();

        //清除临时文件
        ((SXSSFWorkbook)workbook).dispose();

        //记录结束时间
        long end = System.currentTimeMillis();
        System.out.println((double)(end - begin)/1000);

    }






    // 写入 03版本 POI - Excel 读  HSSF
    @Test
    void testRead03() throws Exception{

        InputStream is = new FileInputStream(path + "观众统计表 03.xls");
        Workbook workbook = new HSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);
        // 读取第一行第一列
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);
        // 输出单元内容
        System.out.println(cell.getStringCellValue());
        // 操作结束，关闭文件
        is.close();

    }

    // 写入  POI - Excel 读  XSSF
    @Test
    void testRead07() throws Exception{

        InputStream is = new FileInputStream(path + "观众统计表 07.xlsx");
        Workbook workbook = new XSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);
        // 读取第一行第一列
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);
        // 输出单元内容
        System.out.println(cell.getStringCellValue());
        // 操作结束，关闭文件
        is.close();
    }

    // 写入 03 版本  读取不同的数据类型 HSSF
    @Test
    void testCellType() throws Exception{

        InputStream is = new FileInputStream(path + "会员消费商品明细表.xls");
        Workbook workbook = new HSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);

        // 读取标题所有内容 （Excel表格第一行）
        Row rowTitle = sheet.getRow(0);
        // 行不为空
        if (rowTitle != null) {
            // 读取 第一行列的个数
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = rowTitle.getCell(cellNum);
                if (cell != null) {
                    // 获取 第一行的 字符串数据
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + "|");
                }
            }
            System.out.println();
        }

        // 读取除了第一行之外的表单内容
        // 行的统计，记录
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
            Row rowData = sheet.getRow(rowNum);
            if (rowData != null) {// 行不为空
                // 读取列，这一行有多少列 个数
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    System.out.print("【" + (rowNum + 1) + "-" + (cellNum + 1) + "】");
                    // 得到这个列
                    Cell cell = rowData.getCell(cellNum);
                    if (cell != null) {
                        //进行匹配这个列 的类型
                        int cellType = cell.getCellType();
                        //判断单元格数据类型
                        String cellValue = "";
                        switch (cellType) {
                            case HSSFCell.CELL_TYPE_STRING:  //字符串
                                System.out.print("【STRING】");
                                cellValue = cell.getStringCellValue();
                                break;
                            case HSSFCell.CELL_TYPE_BOOLEAN: //布尔
                                System.out.print("【BOOLEAN】");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case HSSFCell.CELL_TYPE_BLANK:  //空
                                System.out.print("【BLANK】");
                                break;
                            case HSSFCell.CELL_TYPE_NUMERIC: //数字
                                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                    // 日期类型
                                    System.out.print("【日期】");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
                                }else {
                                    // 不是日期格式，则防止当数字过长时以科学计数法显示
                                    System.out.print("【转换成字符串】");
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue = cell.toString();
                                }
                                break;
                            case HSSFCell.CELL_TYPE_ERROR:
                                System.out.print("【数据类型错误】");
                                break;
                        }
                        System.out.println(cellValue);
                    }
                }
            }
        }
        is.close();
    }

    // 写入 07 版本  读取不同的数据类型  XSSF
    @Test
    void testRead07Type() throws Exception{

        InputStream is = new FileInputStream(path + "观众统计表 07.xlsx");
        Workbook workbook = new XSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);

        // 读取标题所有内容 （Excel表格第一行）
        Row rowTitle = sheet.getRow(0);
        // 行不为空
        if (rowTitle != null) {
            // 读取 第一行列的个数
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = rowTitle.getCell(cellNum);
                if (cell != null) {
                    // 获取 第一行的 字符串数据
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + "|");
                }
            }
            System.out.println();
        }

        // 读取除了第一行之外的表单内容
        // 行的统计，记录
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
            Row rowData = sheet.getRow(rowNum);
            if (rowData != null) {// 行不为空
                // 读取列，这一行有多少列 个数
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    System.out.print("【" + (rowNum + 1) + "-" + (cellNum + 1) + "】");
                    // 得到这个列
                    Cell cell = rowData.getCell(cellNum);
                    if (cell != null) {
                        //进行匹配这个列 的类型
                        int cellType = cell.getCellType();
                        //判断单元格数据类型
                        String cellValue = "";
                        switch (cellType) {
                            case XSSFCell.CELL_TYPE_STRING:  //字符串
                                System.out.print("【STRING】");
                                cellValue = cell.getStringCellValue();
                                break;
                            case XSSFCell.CELL_TYPE_BOOLEAN: //布尔
                                System.out.print("【BOOLEAN】");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case XSSFCell.CELL_TYPE_BLANK:  //空
                                System.out.print("【BLANK】");
                                break;
                            case XSSFCell.CELL_TYPE_NUMERIC: //数字
                                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                    // 日期类型
                                    System.out.print("【日期】");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
                                }else {
                                    // 不是日期格式，则防止当数字过长时以科学计数法显示
                                    System.out.print("【转换成字符串】");
                                    cell.setCellType(XSSFCell.CELL_TYPE_STRING);
                                    cellValue = cell.toString();
                                }
                                break;
                            case XSSFCell.CELL_TYPE_ERROR:
                                System.out.print("【数据类型错误】");
                                break;
                        }
                        System.out.println(cellValue);
                    }
                }
            }
        }
        is.close();
    }


    // 计算公式
    @Test
    void testFormula() throws Exception{

        InputStream is = new FileInputStream(path + "计算公式.xls");
        Workbook workbook = new HSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);

        // 读取第五行第一列
        Row row = sheet.getRow(4);
        Cell cell = row.getCell(0);

        //公式计算器
        FormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook) workbook);
        // 输出单元内容
        int cellType = cell.getCellType();
        switch (cellType) {
            case Cell.CELL_TYPE_FORMULA:
                // 2得到公式
                String formula = cell.getCellFormula();
                System.out.println(formula);
                CellValue evaluate = formulaEvaluator.evaluate(cell);
                // String cellValue = String.valueOf(evaluate.getNumberValue());
                String cellValue = evaluate.formatAsString();
                System.out.println(cellValue);
                break;
        }

    }


}
