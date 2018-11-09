package com.lcf.start;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

public class Main {


    public static void main(String[] args) {
        List<String> strings=getFile("D:\\sales");
        int i=0;
        for (String excel_url:strings) {
            try {
                Properties properties = new Properties();

                InputStream in = new BufferedInputStream(new FileInputStream("D:\\sales\\employee.properties"));
                properties.load(new InputStreamReader(in, "utf-8"));
                //获取key对应的value值

                xlsx_reader(excel_url, properties,i++);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    //*************xlsx文件读取函数************************
    //excel_name为文件名，arg为需要查询的列号
    //返回二维字符串数组
    static ArrayList<ArrayList<String>> xlsx_reader(String excel_url, Properties properties,int n) throws IOException {
        FileOutputStream excelFileOutPutStream = new FileOutputStream("D:\\sales\\outPut_"+n+".xls");

        //读取xlsx文件
        HSSFWorkbook HSSfWorkbook = null;
        //寻找目录读取文件
        File excelFile = new File(excel_url);
        InputStream is = new FileInputStream(excelFile);
        HSSfWorkbook = new HSSFWorkbook(is);

        if (HSSfWorkbook == null) {
            System.out.println("未读取到内容,请检查路径！");
            return null;
        }

        ArrayList<ArrayList<String>> ans = new ArrayList<ArrayList<String>>();
        HSSFSheet HSSfSheet = HSSfWorkbook.getSheetAt(0);

        // 对于sheet，读取其中的每一行
        for (int rowNum = 1; rowNum <= HSSfSheet.getLastRowNum(); rowNum++) {
            try {
                HSSFRow HSSfRow = HSSfSheet.getRow(rowNum);
                if (HSSfRow == null) {
                    continue;
                }
                HSSFCell cellA = HSSfRow.getCell(0);
                if (!cellA.getCellTypeEnum().equals(CellType.NUMERIC)) {
                    continue;
                }
                int i = 1;
                HSSFCell cellB = HSSfRow.getCell(i++);
                HSSFCell cellC = HSSfRow.getCell(i++);
                HSSFCell cellD = HSSfRow.getCell(i++);
                HSSFCell cellE = HSSfRow.getCell(i++);
                HSSFCell cellF = HSSfRow.getCell(i++);
                HSSFCell cellG = HSSfRow.getCell(i++);
                HSSFCell cellH = HSSfRow.getCell(i++);
                HSSFCell cellI = HSSfRow.getCell(i++);
                HSSFCell cellJ = HSSfRow.getCell(i++);
                HSSFCell cellK = HSSfRow.getCell(i++);
                HSSFCell cellL = HSSfRow.getCell(i++);
                HSSFCell cellM = HSSfRow.getCell(i++);
                HSSFCell cellN = HSSfRow.getCell(i++);
                HSSFCell cellO = HSSfRow.getCell(i++);
                HSSFCell cellP = HSSfRow.getCell(i++);
                if (cellA.getNumericCellValue() == 1) {
                    HSSFRow row = HSSfSheet.getRow(rowNum - 1);
                    row.createCell(i).setCellValue("提成类型");
                    row.createCell(i + 1).setCellValue("提成金额");
                }

                //销售额
                double sum = cellP.getNumericCellValue();
                //姓名
                String name = cellB.getStringCellValue();

                int type = Integer.parseInt(properties.getProperty(name));
                double tc = getTC(sum, type);

                HSSfRow.createCell(i++).setCellValue(type);
                HSSfRow.createCell(i).setCellValue(tc);
            } catch (Exception e) {
                continue;
            }


        }
        HSSfWorkbook.write(excelFileOutPutStream);
        excelFileOutPutStream.flush();
        excelFileOutPutStream.close();
        return ans;
    }


    //字符串修剪  去除所有空白符号 ， 问号 ， 中文空格
    static private String Trim_str(Object obj) {

        if (obj == null) {
            return null;
        }
        String str = String.valueOf(obj);
        return str.replaceAll("[\\s\\?]", "").replace("　", "");
    }

    static double getTC(double sum, int type) {
        if (type == 1) {
            //提成1
            double tc1 = 0;
            if (sum > 8000 && sum < 25000) {
                tc1 = (sum-8000) * 0.15;
            } else if (sum >= 25000) {
                tc1 = (25000-8000) * 0.15 + (sum - 25000) * 0.2;
            }
            return tc1;
        } else if (type == 2) {
            //提成2
            double tc2 = 0;
            if (sum < 6000) {
                tc2 = sum * 0.05;
            } else if (sum < 15000) {
                tc2 = sum * 0.08;
            } else if (sum < 25000) {
                tc2 = 15000 * 0.08 + (sum - 15000) * 0.1;
            } else if (sum < 35000) {
                tc2 = 15000 * 0.08 + (25000 - 15000) * 0.1 + (sum - 25000) * 0.13;
            } else if (sum >= 35000) {
                tc2 = 15000 * 0.08 + (25000 - 15000) * 0.1 + (35000 - 25000) * 0.13 + (sum - 35000) * 0.2;
            }
            return tc2;
        } else if (type == 3) {
            //提成3
            double tc3 = 0;
            if (sum < 6000) {
                tc3 = sum * 0.05;
            } else if (sum < 20000) {
                tc3 = sum * 0.08;
            } else if (sum < 35000) {
                tc3 = 20000 * 0.08 + (sum - 20000) * 0.1;
            } else if (sum < 50000) {
                tc3 = 20000 * 0.08 + (35000 - 20000) * 0.1 + (sum - 35000) * 0.13;
            } else if (sum >= 50000) {
                tc3 = 20000 * 0.08 + (35000 - 20000) * 0.1 + (50000 - 35000) * 0.13 + (sum - 50000) * 0.2;
            }
            return tc3;
        }else if (type == 4) {//尉团队
            //提成4
            double tc4 = 0;
            if (sum < 15000) {
                tc4 = sum * 0.08;
            } else if (sum < 25000) {
                tc4 = 15000 * 0.08 + (sum - 15000) * 0.1;
            } else if (sum < 40000) {
                tc4 = 15000 * 0.08 + (25000 - 15000) * 0.1 + (sum - 25000) * 0.15;
            } else if (sum >= 40000) {
                tc4 = 15000 * 0.08 + (25000 - 15000) * 0.1 + (40000 - 25000) * 0.15 + (sum - 40000) * 0.2;
            }
            return tc4;
        }
        return 0;


    }

    static List<String> getFile(String directoryPath) {
        List<String> strings=new ArrayList<>();
        List<String> list = new ArrayList<String>();
        File baseFile = new File(directoryPath);

        File[] files = baseFile.listFiles();
        for (File file : files) {
            if (!file.isDirectory()) {
                String name=file.getName();
                String[] names=name.split("\\.");
                if(names.length>=2 &&names[0].contains("职员销售")&&names[names.length-1].equals("xls")){
                    strings.add(directoryPath+"\\"+ name);
                }
            }
        }
        return strings;
    }



}
