package com.im.modules.entities;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

import java.io.Serializable;
import java.util.Date;

/**
 * 用户实体类
 * @Title: Provider
 * @Version: 1.0
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
//@Accessors(chain = true)
public class User implements Serializable {

    @ExcelProperty("ID值")
    private Integer id;
    //用户名
    @ExcelProperty("用户名")
	private String username;
    //真实姓名
    @ExcelProperty("真实姓名")
    private String realName;

    //用户密码
    //Excel忽略这个字段
    @ExcelIgnore
    private String password;

    //性别：1 女  2 男
    @ExcelProperty("性别")
    private Integer gender;
    //生日
    @ExcelProperty("生日")
    private Date birthday;
    //1管理员  2经理  3普通用户
    @ExcelProperty("职称")
    private Integer userType;


}
