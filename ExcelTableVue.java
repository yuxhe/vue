package com.testmaven;

import java.awt.Color;
import java.io.BufferedWriter;
import java.io.File;
/**
 * 算法：
 * 1)扫描出每行 有几个字段  col 分配 24/col   最后一个col=24 - 本行前面所有的col
 * 2)按行扫描出卷标 排好顺序 
 * 3)扫描出类型 按控件类型组装出数据
 *
 */
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/***
 * 
 * @author yuxh
 * 算法:思想        0)预处理  拆分(全选、序号 map)  字段属性(list对应)  按钮属性(字段对应)
 * 1、总共两行数据  1)卷标名称  2)字段名称  按钮识别 
 *
 */

public class ExcelTableVue {
     private static final DataFormatter FORMATTER = new DataFormatter();
     private static String vue_path = "D://呈祥化工201801";
     private static String vuefile = "输出table_vue";
     private static String tab2="\t\t";
     private static String tab3="\t\t\t";
     private static String tab4="\t\t\t\t";
     private static String tab5="\t\t\t\t\t";
     private static String tab6="\t\t\t\t\t\t";
     private static String tab7="\t\t\t\t\t\t\t";
     private static String tab8="\t\t\t\t\t\t\t\t";
     //------------------------登记col列
     private static List<String>  lables=new ArrayList<String>();//卷标
     private static List<String>  fields=new ArrayList<String>();//字段名或 属性
     private static List<String>  colors=new ArrayList<String>();//颜色获取
     
     private static List<String> mapbuttonlabe=new ArrayList<String>();
     private static List<String> mapbuttonstyle=new ArrayList<String>();
     
     private static Map<Integer,String> mapheader=new HashMap<Integer,String>();//登记二维的
     //----------------------------------------------------------------------
     public static void main(String[] args) { 
         //String filePath = "C://Users//yuxhe//Desktop//我的屏幕//table数据.xlsx";
         String filePath = "D://呈祥化工201801//table数据人员信息.xlsx";
         int sheetIndex = 0;
         getExcelValue(filePath, sheetIndex) ;
         System.out.println("---------------");
     }
     //---------------------------------------------------------------------
     /**
      * 获取单元格内容
     * 
      * @param cell
      *            单元格对象
      * @return 转化为字符串的单元格内容
      */
     private static String getCellContent(Cell cell) {
        return FORMATTER.formatCellValue(cell);
     }

    private static String getExcelValue(String filePath, int sheetIndex) {
        String value = "";
         try {
            //创建对Excel工作簿文件
             Workbook book = null;
             try {
               book = new XSSFWorkbook(new FileInputStream(filePath));
          } catch (Exception ex) {
               book = new HSSFWorkbook(new FileInputStream(filePath));
           }

            Sheet sheet = book.getSheetAt(sheetIndex);
           // 获取到Excel文件中的所有行数
           //int rows = sheet.getPhysicalNumberOfRows();
           int rows = sheet.getLastRowNum();
           //----------------------------------------------------
           File folder = new File(vue_path);
           if ( !folder.exists() ) {
               folder.mkdir();
           }
           File filevue = new File(vue_path, vuefile + ".vue");
           //----------------------------------
           FileWriter fw =  new FileWriter(filevue);
           fw.write("");
           fw.close();
           //----------------------------------
           BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(filevue)));
           // 遍历行
           //----------------------------------------------预处理行列
           setPreCol(book,sheet, rows);
           /*
           if   (1==1) {
        	    return  "";
           }
           */
           //----------------------------------------------
           //-------------遍历全选
           int i_break=0;
           for  (Integer keyi:mapheader.keySet()) {
        	     if  ("全选".equals(mapheader.get(keyi))) {
        	    	 bw.write(tab2+"<el-table :data=\"dataRows\" border highlight-current-row v-loading=\"listLoading\"  @selection-change=\"changeAll\" :stripe=\"true\" style=\"width: 100%;\">");
        	    	 i_break=1;
        	    	 break;
        	     }
           }
           //-------------无全选
            if  (i_break==0) {//
                bw.write(tab2+"<el-table :data=\"dataRows\" border highlight-current-row v-loading=\"listLoading\" :stripe=\"true\" style=\"width: 100%;\">");
           }
           //-------------序号是否存在
            for  (Integer keyi:mapheader.keySet()) {
       	       if  ("序号".equals(mapheader.get(keyi))) {
       	    	   bw.newLine();
       	    	   bw.write(tab3+"<el-table-column type=\"index\" header-align=\"center\" align=\"center\" width=\"80\"></el-table-column>");
       	    	 break;
       	       }
           }
           //---------------核查创建文本域
            for  (int i=0;i<colors.size();i++)  {//createColumn(bw,i);
            	switch (colors.get(i)) 
   	    	 { 
   	    	 case "#000000": //el-input text
   	    		 createColumn(bw,i);
   	    		 break;
   	    	case "#969696": //el-input text
  	    		 createInput(bw,i);
  	    		 break;
   	    	 case "#ff5050": //el-input textarea
   	    		 //System.out.println(1);
   	    		 createTextarea(bw,i);
   	    		 break;
   	    	 case "#ff0000": //el-input-number
   	    		 //System.out.println(2);
   	    		 createInputNum(bw,i);
   	    		 break;
   	    	 case "#ffc000": //el-autocomplete
   	    		 System.out.println("el-autocomplete"); 
       	    	 break;
   	    	 case "#92d050": //radio
   	    		 //System.out.println(2);
   	    		 createRadio(bw,i);
       	    	 break; 
   	    	 case "#00b050": //checkbox
   	    		 //System.out.println(2); 
   	    		 createCheckbox(bw,i);
       	    	 break; 
   	    	 case "#00b0f0": //el-switch
   	    		 //System.out.println(2);
   	    		 createSwitch(bw,i);
       	    	 break; 
   	    	 case "#0070c0": //select
   	    		 //System.out.println(2);
   	    		 createSelect(bw,i);
       	    	 break; 
   	    	 case "#2371ff": //el-date-picker
   	    		 //System.out.println(2);
   	    		 createDate(bw,i);
       	    	 break; 
   	    	 case "#7030a0": //el-time-picker
   	    		 //System.out.println(2);
   	    		 createTime(bw,i);
       	    	 break; 
   	    	 case "#cc0000": //el-button
   	    		 System.out.println("el-button"); 
       	    	 break; 
   	    	 case "#ff6699": //span
   	    		 //System.out.println(2);
   	    		 createSpan(bw,i);
       	    	 break;
   	    	 default: 
   	    	     ;
   	    	     break; 
   	    	 }
            	
            }
           //----------------------------------------------
            for  (int i=0;i<mapbuttonlabe.size();i++)  {
                 createButton(bw,i);
            }
            //--------------创建按钮
            bw.newLine();
 		    bw.write(tab2+"</el-table>");
            bw.flush();//数据文件流关闭了
            bw.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
         } catch (IOException e) {
            e.printStackTrace();
      }

        return value;

     }

   
    
   /***
    * 颜色转换
    * @param color
    * @return
    */
   public static String Color2String(Color color) {  
        String R = Integer.toHexString(color.getRed());  
        R = R.length() < 2 ? ('0' + R) : R;  
        String G = Integer.toHexString(color.getGreen());  
        G = G.length() < 2 ? ('0' + G) : G;  
        String B = Integer.toHexString(color.getBlue());  
        B = B.length() < 2 ? ('0' + B) : B;  
        return '#' + R + G + B;  
    } 
   
   /***
    * 处理创建单行编辑框
    * @param bw  sheet  i  j colcur colcnt
    * @throws IOException
    */
   public static void  createColumn(BufferedWriter bw,int i) throws IOException {//普通input text
	   //获取卷标名称
	   String  labelStr=lables.get(i);
	   String  field=fields.get(i); //字段名称
       String[] splits=field.split("\n");//回车换行
	   if  (splits.length >1) {//
		   bw.newLine();
		   bw.write(tab3+"<el-table-column prop=\""+splits[0]+"\" label=\""+labelStr+"\" header-align=\"center\" align=\"center\">");
		   bw.newLine();
		   bw.write(tab4+"<template slot-scope=\"scope\">");
		   for (int k=1;k<splits.length;k++) {
			   String[]  str=splits[k].split("#");
			   bw.newLine();
			   bw.write(tab5+"<span v-if=\"scope.row."+splits[0]+" ==\'"+str[0]+"\'>"+str[1]+"</span>");
		   }
		   bw.newLine();
		   bw.write(tab4+"</template>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
	   }else {//普通格式的
		   bw.newLine();
		   bw.write(tab3+"<el-table-column prop=\""+splits[0]+"\" label=\""+labelStr+"\" header-align=\"center\" align=\"center\"></el-table-column>");		   
	   }
   }
   
   /***
    *   可编辑类的创建
    */
    //----------------------------------------------------------------------
   /***创建可编辑的 单行编辑框
    * @param bw
    * @param i
    * @throws IOException
    */
   public static void  createInput(BufferedWriter bw,int i) throws IOException {//普通input text
	   //获取卷标名称
	   String  labelStr=lables.get(i);
	   String  field=fields.get(i); //字段名称
       String[] splits=field.split("\n");//回车换行
		   bw.newLine();
		   bw.write(tab3+"<el-table-column prop=\""+splits[0]+"\" label=\""+labelStr+"\" header-align=\"center\" align=\"center\">");
		   bw.newLine();
		   bw.write(tab4+"<template slot-scope=\"scope\">");
		   bw.newLine();
		   bw.write(tab5+"<el-input  ");
		   for (int k=1;k<splits.length;k++) {
			   bw.write("  " + splits[k] +"  " );
		   }
		   bw.write("  v-model=\"scope.row."+splits[0]+"\"  placeholder=\"请输入"+labelStr+"\"></el-input>");
		   bw.newLine();
		   bw.write(tab4+"</template>");
		   bw.newLine();
		   bw.write(tab4+"</el-table-column>");
   }
   
   
   /***创建可编辑的 多行编辑框
    * @param bw
    * @param i
    * @throws IOException
    */
   public static void  createTextarea(BufferedWriter bw,int i) throws IOException {//普通input text
	   //获取卷标名称
	   String  labelStr=lables.get(i);
	   String  field=fields.get(i); //字段名称
       String[] splits=field.split("\n");//回车换行
		   bw.newLine();
		   bw.write(tab3+"<el-table-column prop=\""+splits[0]+"\" label=\""+labelStr+"\" header-align=\"center\" align=\"center\">");
		   bw.newLine();
		   bw.write(tab4+"<template slot-scope=\"scope\">");
		   bw.newLine();
		   bw.write(tab5+"<el-input  type=\"textarea\"  ");
		   for (int k=1;k<splits.length;k++) {
			   bw.write("  " + splits[k] +"  " );
		   }
		   bw.write("  v-model=\"scope.row."+splits[0]+"\"  placeholder=\"请输入"+labelStr+"\"></el-input>");
		   bw.newLine();
		   bw.write(tab4+"</template>");
		   bw.newLine();
		   bw.write(tab4+"</el-table-column>");
   }
   
   /***创建可编辑的 多行编辑框
    * @param bw
    * @param i
    * @throws IOException
    */
   public static void  createInputNum(BufferedWriter bw,int i) throws IOException {//
	   //获取卷标名称
	   String  labelStr=lables.get(i);
	   String  field=fields.get(i); //字段名称
       String[] splits=field.split("\n");//回车换行
		   bw.newLine();
		   bw.write(tab3+"<el-table-column prop=\""+splits[0]+"\" label=\""+labelStr+"\" header-align=\"center\" align=\"center\">");
		   bw.newLine();
		   bw.write(tab4+"<template slot-scope=\"scope\">");
		   bw.newLine();
		   bw.write(tab5+"<el-input-number   ");
		   for (int k=1;k<splits.length;k++) {
			   bw.write("  " + splits[k] +"  " );
		   }
		   bw.write("  v-model=\"scope.row."+splits[0]+"\"  placeholder=\"请输入"+labelStr+"\"  :min=\"1\" :max=\"10\"></el-input-number>");
		   bw.newLine();
		   bw.write(tab4+"</template>");
		   bw.newLine();
		   bw.write(tab4+"</el-table-column>");
   }
   
   /***
    * 创建单选按钮组哈
    * @param bw
    * @param i
    * @throws IOException
    */
   public static void  createRadio(BufferedWriter bw,int i) throws IOException {//
		   //获取卷标名称
		   String  labelStr=lables.get(i);
		   String  field=fields.get(i); //字段名称
	       String[] splits=field.split("\n");//回车换行
       
		   bw.newLine();
		   bw.write(tab3+"<el-table-column prop=\""+splits[0]+"\" label=\""+labelStr+"\" header-align=\"center\" align=\"center\">");
		   bw.newLine();
		   bw.write(tab4+"<template slot-scope=\"scope\">");
		   bw.newLine();
		   bw.write(tab5+"<el-radio-group  v-model=\"scope.row."+splits[0]+"\"  >");
		   if  (field.indexOf("#") >=0) {//单选按钮
			   for  (int k=1;k<splits.length;k++) {
				     String[]  strs=splits[k].split("#");
				     bw.newLine();
				     bw.write(tab6+"<el-radio  "+ strs[0]+">"+strs[1]+"</el-radio>");
			   }
		   }else { //radio in 方式
			   bw.newLine();
			   bw.write(tab6+"<el-radio ");
			   for  (int k=1;k<splits.length;k++) {
				     bw.newLine();
				     bw.write(tab7+ splits[k]);
			   }
			   bw.newLine();
			   bw.write(tab7+">");
			   bw.newLine();
			   bw.write(tab6+"</el-radio>");
		   }
		   bw.newLine();
		   bw.write(tab5+"</el-radio-group>");

		   bw.newLine();
		   bw.write(tab4+"</template>");
		   bw.newLine();
		   bw.write(tab4+"</el-table-column>"); 
   }
   
   /***
    * 创建单选按钮组哈
    * @param bw
    * @param i
    * @throws IOException
    */
   public static void  createCheckbox(BufferedWriter bw,int i) throws IOException {//
		   //获取卷标名称
		   String  labelStr=lables.get(i);
		   String  field=fields.get(i); //字段名称
	       String[] splits=field.split("\n");//回车换行
       
		   bw.newLine();
		   bw.write(tab3+"<el-table-column prop=\""+splits[0]+"\" label=\""+labelStr+"\" header-align=\"center\" align=\"center\">");
		   bw.newLine();
		   bw.write(tab4+"<template slot-scope=\"scope\">");
		   bw.newLine();
		   bw.write(tab5+"<el-checkbox-group  v-model=\"scope.row."+splits[0]+"\"  >");
		   if  (field.indexOf("#") >=0) {//单选按钮
			   for  (int k=1;k<splits.length;k++) {
				     String[]  strs=splits[k].split("#");
				     bw.newLine();
				     bw.write(tab6+"<el-checkbox   "+ strs[0]+">"+strs[1]+"</el-checkbox>");
			   }
		   }else { //radio in 方式
			   bw.newLine();
			   bw.write(tab6+"<el-checkbox ");
			   for  (int k=1;k<splits.length;k++) {
				     bw.newLine();
				     bw.write(tab7+ splits[k]);
			   }
			   bw.newLine();
			   bw.write(tab7+">");
			   bw.newLine();
			   bw.write(tab8+"</el-checkbox>");
		   }
		   bw.newLine();
		   bw.write(tab5+"</el-checkbox-group>");

		   bw.newLine();
		   bw.write(tab4+"</template>");
		   bw.newLine();
		   bw.write(tab4+"</el-table-column>"); 
   }
   
   /***
    * 创建 下拉选择框
    * @param bw
    * @param i
    * @throws IOException
    */
   public static void  createSelect(BufferedWriter bw,int i) throws IOException {//
		   //获取卷标名称
		   String  labelStr=lables.get(i);
		   String  field=fields.get(i); //字段名称
	       String[] splits=field.split("\n");//回车换行
       
		   bw.newLine();
		   bw.write(tab3+"<el-table-column prop=\""+splits[0]+"\" label=\""+labelStr+"\" header-align=\"center\" align=\"center\">");
		   bw.newLine();
		   bw.write(tab4+"<template slot-scope=\"scope\">");
		   bw.newLine();
		   bw.write(tab5+"<el-select  v-model=\"scope.row."+splits[0]+"\"   >");
		   if  (field.indexOf("#") >=0) {//单选按钮
			   for  (int k=1;k<splits.length;k++) {
				     String[]  strs=splits[k].split("#");
				     bw.newLine();
				     bw.write(tab6+"<el-option  "+ strs[0]+">"+strs[1]+"</el-option>");
			   }
		   }else { //radio in 方式
			   bw.newLine();
			   bw.write(tab6+"<el-option ");
			   for  (int k=1;k<splits.length;k++) {
				     bw.newLine();
				     bw.write(tab7+ splits[k]);
			   }
			   bw.newLine();
			   bw.write(tab7+">");
			   bw.newLine();
			   bw.write(tab6+"</el-option>");
		   }
		   bw.newLine();
		   bw.write(tab5+"</el-select>");

		   bw.newLine();
		   bw.write(tab4+"</template>");
		   bw.newLine();
		   bw.write(tab4+"</el-table-column>"); 
   }
   
   /***
    * 创建 下拉选择框
    * @param bw
    * @param i
    * @throws IOException
    */
   public static void  createSwitch(BufferedWriter bw,int i) throws IOException {//
		   //获取卷标名称
		   String  labelStr=lables.get(i);
		   String  field=fields.get(i); //字段名称
	       String[] splits=field.split("\n");//回车换行
       
		   bw.newLine();
		   bw.write(tab3+"<el-table-column prop=\""+splits[0]+"\" label=\""+labelStr+"\" header-align=\"center\" align=\"center\">");
		   bw.newLine();
		   bw.write(tab4+"<template slot-scope=\"scope\">");
		   bw.newLine();
		   
		   bw.write(tab5+"<el-switch  v-model=\"scope.row."+splits[0]+"\"  ");
		   for  (int k=1;k<splits.length;k++) {
			     bw.newLine();
			     bw.write(tab6+ splits[k]);
		   }
		   bw.newLine();
		   bw.write(tab7+">");
		   bw.newLine();
		   bw.write(tab8+"</el-switch>");
		   
		   bw.newLine();
		   bw.write(tab4+"</template>");
		   bw.newLine();
		   bw.write(tab4+"</el-table-column>"); 
   }
   
   /***
    * 处理创建date
    * @param bw  sheet  i  j colcur colcnt
    * @throws IOException
    */
   public static void  createDate(BufferedWriter bw,int i) throws IOException {//普通input text
	   //获取卷标名称
	   String  labelStr=lables.get(i);
	   String  field=fields.get(i); //字段名称
       String[] splits=field.split("\n");//回车换行

       bw.newLine();
	   bw.write(tab3+"<el-table-column prop=\""+splits[0]+"\" label=\""+labelStr+"\" header-align=\"center\" align=\"center\">");
	   bw.newLine();
	   bw.write(tab4+"<template slot-scope=\"scope\">");
	   bw.newLine();
	   bw.write(tab5+"<el-date-picker type=\"date\" v-model=\"scope.row."+splits[0]+"\"  placeholder=\"请选择日期"+"\"   style=\"width: 100%;\"></el-date-picker>");
	   bw.newLine();
	   bw.write(tab4+"</template>");
	   bw.newLine();
	   bw.write(tab3+"</el-table-column>");
   }
   
   
   /***
    * 处理创建date
    * @param bw  sheet  i  j colcur colcnt
    * @throws IOException
    */
   public static void  createTime(BufferedWriter bw,int i) throws IOException {//普通input text
	   //获取卷标名称
	   String  labelStr=lables.get(i);
	   String  field=fields.get(i); //字段名称
       String[] splits=field.split("\n");//回车换行

       bw.newLine();
	   bw.write(tab3+"<el-table-column prop=\""+splits[0]+"\" label=\""+labelStr+"\" header-align=\"center\" align=\"center\">");
	   bw.newLine();
	   bw.write(tab4+"<template slot-scope=\"scope\">");
	   bw.newLine();
       
	   bw.write(tab5+"<el-time-picker type=\"fixed-time\" v-model=\"scope.row."+splits[0]+"\"  placeholder=\"请选择日期"+"\"   style=\"width: 100%;\"></el-time-picker>");
	   bw.newLine();
	   bw.write(tab4+"</template>");
	   bw.newLine();
	   bw.write(tab3+"</el-table-column>");
   }
  
   /***
    * 处理创建Span  
    * @param bw  sheet  i  j colcur colcnt
    * @throws IOException
    */
   public static void  createSpan(BufferedWriter bw,int i) throws IOException {//
	   //获取卷标名称
	   String  labelStr=lables.get(i);
	   String  field=fields.get(i); //字段名称
       String[] splits=field.split("\n");//回车换行
		   bw.newLine();
		   bw.write(tab3+"<el-table-column prop=\""+splits[0]+"\" label=\""+labelStr+"\" header-align=\"center\" align=\"center\">");
		   bw.newLine();
		   bw.write(tab4+"<template slot-scope=\"scope\">");
		   bw.newLine();
		   bw.write(tab5+"<span  ");
		   for (int k=1;k<splits.length;k++) {
			   bw.write("  " + splits[k] +"  " );
		   }
		   
		   bw.newLine();
		   bw.write(" ></span>");
		   bw.newLine();
		   bw.write(tab4+"</template>");
		   bw.newLine();
		   bw.write(tab4+"</el-table-column>");
   }
   
   //-----------------------------------------------------------------------
   
   
   /***
    * 创建序号
    * @param bw
    * @throws IOException
    */
   public static void  createButton(BufferedWriter bw,int i) throws IOException {//
	   
	   String  labelStr=mapbuttonlabe.get(i);
	   String  field=mapbuttonstyle.get(i); 
       if (i==0) {
    	   bw.newLine();
    	   bw.write(tab3+"<el-table-column label=\"操作\" width=\"150\" header-align=\"center\">");
    	   bw.newLine();
    	   bw.write(tab4+"<template slot-scope=\"scope\">");
       }
       bw.newLine();
       bw.write(tab5+"<el-button size=\""+field+"\" @click=\"handleEdit(scope.$index, scope.row)\">"+labelStr+"</el-button>");

	   if (i==mapbuttonlabe.size() - 1) {//已到末尾
		   bw.newLine();
		   bw.write(tab4+"</template>");
		   bw.newLine();
		   bw.write(tab3+"</el-table-column>");
	   }   }
   
     
   /***
    * 预处理数据项
    * @param sheet
    * @param rows
    */
   public  static void setPreCol(Workbook book,Sheet sheet,int rows) {
	   //----------------------------------------------
	   int  i_hh=0;
	   int  row_start=5;
	   //-----------------------确定其实行数据
	   lableA:
	   for (int i = 5; i <= rows; i++) {
		   Row row = sheet.getRow(i);
    	   if (row != null) {
    		   int cells = row.getLastCellNum(); //注意此处应该采用这个东东
               for (int j = 0; j < cells; j++) {
                  // 获取到列的值
                  Cell cell = row.getCell(j);
                  if (cell != null &&  getCellContent(cell)!=null && !"".equals(getCellContent(cell)) ) {
                	  row_start=i;
                	  break lableA; //已确定起行
                  }
               }
    	   }
       }
	   //-----------------------
       for (int i = row_start; i <= rows; i++) {//此处 注意了  rows 为行索引号 从5 行开始 哈
    	   Row row = sheet.getRow(i);
    	   if (row != null) {
                // 获取到Excel文件中的所有的列
    		    i_hh=i_hh +1 ;
                //int cells = row.getPhysicalNumberOfCells();
                //System.out.println(row.getLastCellNum()); 为 列号加一
                int cells = row.getLastCellNum(); //注意此处应该采用这个东东
                for (int j = 0; j < cells; j++) {
                   // 获取到列的值
                   Cell cell = row.getCell(j);
                   
                   if (cell != null &&  getCellContent(cell)!=null && !"".equals(getCellContent(cell)) ) {
                	   
                	   if (i_hh==1 && "全选".equals(getCellContent(cell))) {
                		   mapheader.put(j,"全选");
                		   continue;
                	   }else if (i_hh==1 && "序号".equals(getCellContent(cell))) {
                    	   mapheader.put(j,"序号");
                    	   continue;
                       }
                	   
                	   if  ("全选".equals(getCellContent(cell)) || "序号".equals(getCellContent(cell))) {
                		   continue;
                	   }
                	   
                	   
                	   if (i_hh==2 && mapheader.get(j)!=null) {//视其为按钮哈
                    	   mapbuttonstyle.add(getCellContent(cell)); //按钮颜色样式描述 
                    	   continue;
                       }

                	   //-----------------------------------处理按钮
                	   Font eFont = book.getFontAt(cell.getCellStyle().getFontIndex());
	          	       XSSFFont f = (XSSFFont) eFont;
	          	       byte[] rgb=f.getXSSFColor().getRGB();
	          	       Color bakcolor=new java.awt.Color(rgb[0] & 0xFF, rgb[1] & 0xFF , rgb[2] & 0xFF) ;
	          	       String fontcolor=Color2String(bakcolor);  
	          	       if  ("#cc0000".equals(fontcolor)) {
	          	    	    mapheader.put(j,getCellContent(cell));//登记按钮
	          	    	    mapbuttonlabe.add(getCellContent(cell));//加入按钮数组
	          	    	    continue;
	          	       }
	          	       
                	   //-----------------------------------
                	   //mapheader    0,8 文字  1,8 核查0,8 是否存在 按钮 0,全选  0,序号，（0,取消 非前两者则是 按钮,放入数组)
		          	   System.out.println(i_hh+":"+getCellContent(cell));
	          	       if  (i_hh==1) {
                		   lables.add(getCellContent(cell));//卷标
                		   //------------------------------获取字体颜色
                		   colors.add(fontcolor);
      	          	     System.out.println(getCellContent(cell));
                		   //------------------------------
                	   }else if (i_hh==2) {
                		   fields.add(getCellContent(cell));//字段名
                	   } 
	          	       
                   }  
                }
                
    	   }
       }
       //----------------------------------------------
       System.out.println(mapheader);
       System.out.println(lables);
       System.out.println(colors);
       System.out.println(fields);
       System.out.println(mapbuttonstyle);
       
   } 
}