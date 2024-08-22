package com.maantic.automation.utils;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class Constants {
//    public static final String TEST_DATA_SHEET_PATH= System.getProperty("user.dir")+"\\src\\main\\resources\\exceltestdata\\GBT.xlsx";
    public static final String TEST_DATA_SHEET_PATH= "D:\\GBT\\GBT.xlsx";
    public static final String TEST_OUT_DATA_SHEET_PATH = "D:\\GBT\\GBT-Output\\"+ "GBT_Results_"+outFileName()+".xlsx";
//    public static final String TEST_OUT_DATA_SHEET_PATH= System.getProperty("user.dir")+"\\src\\main\\resources\\output\\GBT_Results_"+outFileName()+".xlsx";
    public static final String EXCEL_SHEET_NAME = "GBT";
    public static final int IMPLICIT_WAIT_TIME= 30;
    public static final int EXPLICIT_WAIT_TIME= 30;
    public static final int PAGE_LOAD_WAIT_TIME= 30;

    public static String outFileName(){
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd-HH-mm-ss");
        LocalDateTime now = LocalDateTime.now();
        System.out.println(dtf.format(now));
        return dtf.format(now).toString();
    }
}
