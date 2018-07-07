package com.sinosoft.service;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.util.StringUtil;
import org.junit.Test;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelListUtil {
    public static void main(String[] args) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(new File("D:\\sourcecode\\idea\\e3\\excel-service\\src\\main\\resources\\数据.xls")));
        Map<String,HSSFSheet> sheetMap = new HashMap<String, HSSFSheet>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {// 获取每个Sheet表
            sheetMap.put(workbook.getSheetAt(i).getSheetName(),workbook.getSheetAt(i));
        }
        List<List<Map<String, Object>>> list = getFinalList(sheetMap,"list");
        System.out.println("final_list="+list);
    }

    public static List<List<Map<String, Object>>> getFinalList(Map<String, HSSFSheet> sheetMap,String sheetName){
        List<List<Map<String, Object>>> list = sheet2List(sheetMap.get(sheetName));
        for (List<Map<String, Object>> childList:list) {
            String pid = "";
            String childSheetName = "";
            for (Map<String,Object> map:childList) {
                for(Map.Entry<String,Object> entry:map.entrySet()){
                    if("id".equals(entry.getKey())){
                        pid = entry.getValue()+"";
                    }
                    if(entry.getValue().toString().contains("#")){
                        childSheetName = entry.getValue().toString().replace("#","");
                    }
                }
            }
            System.err.println("pid="+pid+" childSheetName="+childSheetName);
            List<List<Map<String, Object>>> list2 = sheet2List(sheetMap.get(childSheetName),pid);
            System.out.println("list2="+list2);
            for (Map<String,Object> map:childList) {
                for(Map.Entry<String,Object> entry:map.entrySet()){
                    if(entry.getValue().toString().contains("#")){
                        entry.setValue(list2);
                    }
                }
            }
        }
        return list;
    }

    public static List<List<Map<String,Object>>> sheet2List(HSSFSheet sheet){
        return sheet2List(sheet,"");
    }

    public static List<List<Map<String,Object>>> sheet2List(HSSFSheet sheet, String pid){
        //标识符，如果是循环父列表，则将flag置为true，跳过pid的检测
        boolean flag = false;
        if("".equals(pid)) flag = true;
        List<List<Map<String,Object>>> list = new ArrayList<List<Map<String,Object>>>();
        List<String> keyList = new ArrayList<String>();
        for (int j = 0; j < sheet.getLastRowNum() + 1; j++) {// getLastRowNum，获取最后一行的行标
            List<Map<String,Object>> childList = new ArrayList<Map<String, Object>>();
            HSSFRow row = sheet.getRow(j);
            if(flag||pid.equals(row.getCell(1).toString())||"pid".equals(row.getCell(1).toString())){
                if (row != null) {
                    for (int k = 0; k < row.getLastCellNum(); k++) {// getLastCellNum，是获取最后一个不为空的列是第几个
                        if (row.getCell(k) != null) { // getCell 获取单元格数据
                            if(j==0) {
                                keyList.add(row.getCell(k) + "");
                            }else{
                                HashMap<String, Object> hashMap = new HashMap<String, Object>();
                                hashMap.put(keyList.get(k), row.getCell(k) + "");
                                childList.add(hashMap);
                            }
                        }
                    }
                    if(j!=0) list.add(childList);
                }
            }
        }
        return list;
    }
}
