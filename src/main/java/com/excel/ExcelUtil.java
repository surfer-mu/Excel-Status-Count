package com.excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ExcelUtil {

    private String filePath;
    private List<String[]> list = new ArrayList<String[]>();

    public ExcelUtil(String filePath) {
        this.filePath = filePath;
    }

    private void readExcel() throws IOException, BiffException {
        //创建输入流
        InputStream stream = new FileInputStream(filePath);
        //获取Excel文件对象
        Workbook rwb = Workbook.getWorkbook(stream);
        //获取文件的指定工作表 默认的第一个
        Sheet sheet = rwb.getSheet(0);
        //行数(表头的目录不需要，从1开始)
        for (int i = 0; i < sheet.getRows(); i++) {
            //创建一个数组 用来存储每一列的值
            String[] str = new String[sheet.getColumns()];
            Cell cell = null;
            //列数
            for (int j = 0; j < sheet.getColumns(); j++) {
                //获取第i行，第j列的值
                cell = sheet.getCell(j, i);
                str[j] = cell.getContents();
            }
            //把刚获取的列存入list
            list.add(str);
        }
    }

    /**
     * 输出整张表
     */
    private void outData() {
        for (int i = 0; i < list.size(); i++) {
            String[] str = (String[]) list.get(i);
            for (int j = 0; j < str.length; j++) {
                System.out.print(str[j] + '\t');
            }
            System.out.println();
        }
    }

    /**
     * 输出整张表
     */
    private TreeMap<String, String> countPrice() {
        TreeMap<String, String> treeMap = new TreeMap<String, String>();
        list.remove(0);
        for (int i = 0; i < list.size(); i++) {
            String[] str = list.get(i);
            String count = "";
            int zero = 0,one=0;
            if (treeMap.containsKey(str[3])) {
                count = treeMap.get(str[3]);

                String[] split = count.split("_");
                if(Integer.parseInt(str[4])==0){
                    zero = Integer.parseInt(split[0]) + 1;
                    one = Integer.parseInt(split[1]);
                }
                else if(Integer.parseInt(str[4])==1){
                    zero = Integer.parseInt(split[0]);
                    one = Integer.parseInt(split[1]) + 1;
                }
                count=zero+"_"+one;

            }
            else{
                if(Integer.parseInt(str[4])==0){
                    zero =1;
                }
                else if(Integer.parseInt(str[4])==1){
                    one = 1;
                }
                count=zero+"_"+one;
            }
            treeMap.put(str[3], count);

        }
        return treeMap;
    }

    public static void main(String args[]) throws BiffException, IOException {
        ExcelUtil excel = new ExcelUtil("E:\\source\\IdeaProjects\\Myexcel\\src\\main\\resources\\info.xls");
        excel.readExcel();
        //excel.outData();
        TreeMap<String, String> treeMap = excel.countPrice();
        System.out.println("共计有不同价格：" + treeMap.size());
        for (Map.Entry<String, String> map : treeMap.entrySet()) {
            String[] split = map.getValue().split("_");
            int zero = Integer.parseInt(split[0]);
            int one = Integer.parseInt(split[1]);
            int sum =  zero+ one;
            System.out.print("价格：" + map.getKey() + "，出现次数：" + sum);
            System.out.println("。其中为0时，次数有："+zero+";为1时，次数有："+one);
        }
    }

}