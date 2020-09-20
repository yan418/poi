package com.im.modules.entities;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

import java.io.Serializable;
import java.util.Date;

@Data
@AllArgsConstructor
@NoArgsConstructor
//@Accessors(chain = true)
public class DemoData implements Serializable {

    private String string;
    private Date date;
    private Double doubleData;

}
