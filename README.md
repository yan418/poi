# POI 和 easyExcel
开发中经常会设计到excel的处理，如导出Excel，导入Excel到数据库中！ <br>
操作Excel目前比较流行的就是 Apache POI 和 阿里巴巴的 easyExcel ！<br>

## Apache POI
Apache POI 官网：https://poi.apache.org <br>

``` bash
1.导包
    <!-- xls(03)-->
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi</artifactId>
        <version>3.9</version>
    </dependency>
    <!--xlsx(07)-->
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
        <version>3.9</version>
    </dependency>

    <!--日期格式化工具-->
    <dependency>
        <groupId>joda-time</groupId>
        <artifactId>joda-time</artifactId>
        <version>2.10.1</version>
    </dependency>
    
2.具体使用在测试类查看
```

## easyExcel
easyExcel 官网地址：https://github.com/alibaba/easyexcel <br>

``` bash
导包
    <!-- alibaba 的 EasyExcel - Excel -->
    <dependency>
        <groupId>com.alibaba</groupId>
        <artifactId>easyexcel</artifactId>
        <version>2.1.7</version>
    </dependency>
    <dependency>
        <groupId>com.alibaba</groupId>
        <artifactId>fastjson</artifactId>
        <version>1.2.62</version>
    </dependency>
    
```
