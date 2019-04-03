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
import java.util.HashMap;
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
 *
 */

public class ExcelVue {
     private static final DataFormatter FORMATTER = new DataFormatter();
     private static String vue_path = "D://呈祥化工201801";
     private static String vuefile = "输出xls";
     private static String tab2="\t\t";
     private static String tab3="\t\t\t";
     private static String tab4="\t\t\t\t";
     private static String tab5="\t\t\t\t\t";
     private static String tab6="\t\t\t\t\t\t";
     private static String tab7="\t\t\t\t\t\t\t";
     private static String tab8="\t\t\t\t\t\t\t\t";
     private static String tab9="\t\t\t\t\t\t\t\t\t";
     private static String tab10="\t\t\t\t\t\t\t\t\t\t";
     private static String tab11="\t\t\t\t\t\t\t\t\t\t\t";
     //------------------------登记col列
     private static Map<Integer,Integer>  mapColI=new HashMap<Integer,Integer>();//记录每行所占列数
     
     private static Map<String,Integer>  mapColIcur=new HashMap<String,Integer>();//记录每个二维表 每个格子所占列数
     
     //------------------------登记每行的卷标数
     private static Map<String,String>  mapColLable=new HashMap<String,String>();//记录卷标 按二维表记录
     //----------------------------------------------------------------------
     public static void main(String[] args) { 
         //String filePath = "C://Users//yuxhe//Desktop//我的屏幕//vue字典表.xlsx";
         String filePath = "D://呈祥化工201801//合同仓储费用.xlsx";
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
           int rows = sheet.getPhysicalNumberOfRows();
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
           bw.write(tab2+"<el-form  :model=\"editForm\" :rules=\"editFormRules\" ref=\"editForm\" label-width=\"80px\">");
           //----------------------------------------------------
           // 遍历行
           //----------------------------------------------预处理行列
           setPreCol(sheet, rows);
           //----------------------------------------------
            for (int i = 5; i < rows; i++) {//从第6行开始  rows
                // 读取左上端单元格
                 Row row = sheet.getRow(i);
               // 行不为空
                if (row != null) {
                   // 获取到Excel文件中的所有的列
                   int cells = row.getPhysicalNumberOfCells();
                   // 遍历列
                    int  colcur=0;//当前列
                    int  colcnt=0;//字段列
                    for (int j = 0; j < cells; j++) {
                       // 获取到列的值
                       Cell cell = row.getCell(j);
                       if (cell != null &&  getCellContent(cell)!=null && !"".equals(getCellContent(cell)) ) {
                    	    
                    	    if  (cell.getCellStyle().getFillForegroundColorColor()!=null)  {//是标签继续处理
                    	    	 continue;
                    	    }else {
                    	    	 colcur=colcur +1;
                    	    	 Font eFont = book.getFontAt(cell.getCellStyle().getFontIndex());
                    	    	 XSSFFont f = (XSSFFont) eFont;
                    	    	 byte[] rgb=f.getXSSFColor().getRGB();
                    	    	 Color bakcolor=new java.awt.Color(rgb[0] & 0xFF, rgb[1] & 0xFF , rgb[2] & 0xFF) ;
                    	    	 String fontcolor=Color2String(bakcolor);  
                    	    	 if  (!"#ff6699".equals(fontcolor)) {
                    	    		 colcnt=colcnt+1; //记录列数
                    	    	 }
                    	    	 switch (fontcolor) 
                    	    	 { 
                    	    	 case "#000000": //el-input text
                    	    		 createInput(bw,sheet,i,j,colcur,colcnt);
                    	    		 break;
                    	    	 case "#ff5050": //el-input textarea
                    	    		 //System.out.println(1);
                    	    		 createTextarea(bw,sheet,i,j,colcur,colcnt);
                    	    		 break;
                    	    	 case "#ff0000": //el-input-number
                    	    		 //System.out.println(2);
                    	    		 createInputNum(bw,sheet,i,j,colcur,colcnt);
                    	    		 break;
                    	    	 case "#ffc000": //el-autocomplete
                    	    		 System.out.println("el-autocomplete"); 
                        	    	 break;
                    	    	 case "#92d050": //radio
                    	    		 //System.out.println(2);
                    	    		 createRadio(bw,sheet,i,j,colcur,colcnt);
                        	    	 break; 
                    	    	 case "#00b050": //checkbox
                    	    		 //System.out.println(2); 
                    	    		 createCheckbox(bw,sheet,i,j,colcur,colcnt);
                        	    	 break; 
                    	    	 case "#00b0f0": //el-switch
                    	    		 //System.out.println(2);
                    	    		 createSwitch(bw,sheet,i,j,colcur,colcnt);
                        	    	 break; 
                    	    	 case "#0070c0": //select
                    	    		 //System.out.println(2);
                    	    		 createSelect(bw,sheet,i,j,colcur,colcnt);
                        	    	 break; 
                    	    	 case "#2371ff": //el-date-picker
                    	    		 //System.out.println(2);
                    	    		 createDate(bw,sheet,i,j,colcur,colcnt);
                        	    	 break; 
                    	    	 case "#7030a0": //el-time-picker
                    	    		 //System.out.println(2);
                    	    		 createTime(bw,sheet,i,j,colcur,colcnt);
                        	    	 break; 
                    	    	 case "#cc0000": //el-button
                    	    		 
                    	    		 System.out.println("el-button"); 
                        	    	 break; 
                    	    	 case "#ff6699": //span
                    	    		 //System.out.println(2);
                    	    		 createSpan(bw,sheet,i,j,colcur,colcnt);
                        	    	 break;
                    	    	 default: 
                    	    	     ;
                    	    	     break; 
                    	    	 }
                    	    	 
                    	    }
                             
                    }
                   }

                 }
             }
            bw.newLine();
 		    bw.write(tab2+"</el-form>");
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
   public static void  createInput(BufferedWriter bw,Sheet sheet,int i,int j,int colcur,int colcnt) throws IOException {//普通input text
	   //获取卷标名称
	   String  labelStr=mapColLable.get(i+","+colcnt);
	   Row row = sheet.getRow(i);
	   Cell cell = row.getCell(j);
	   String  field=getCellContent(cell); //字段名称
	   if (colcur==1) {
		   bw.newLine();
		   bw.write(tab3+"<el-row>");
		   bw.newLine();
		   bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">");   //:span="8"
		   bw.newLine();
		   bw.write(tab5+"<el-form-item  label=\""+labelStr+"\"  prop=\""+field+"\">");
		   bw.newLine();
		   bw.write(tab6+"<el-input  v-model=\"editForm."+field+"\"  placeholder=\"请输入"+labelStr+"\"   :maxlength=\"100\"></el-input>");
		   bw.newLine();
		   bw.write(tab5+"</el-form-item>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
	   }else {
		   bw.newLine();
		   bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">");   //:span="8"
		   bw.newLine();
		   bw.write(tab5+"<el-form-item  label=\""+labelStr+"\"  prop=\""+field+"\">");
		   bw.newLine();
		   bw.write(tab6+"<el-input  v-model=\"editForm."+field+"\"  placeholder=\"请输入"+labelStr+"\"   :maxlength=\"100\"></el-input>");
		   bw.newLine();
		   bw.write(tab5+"</el-form-item>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
	   }
	   
	   if  (colcur==mapColI.get(i) || mapColI.get(i)==null) {
		   bw.newLine();
		   bw.write(tab3+"</el-row>");
	   }
   }
   
   
   /***
    * 处理创建多行编辑框
    * @param bw  sheet  i  j colcur colcnt
    * @throws IOException
    */
   public static void  createTextarea(BufferedWriter bw,Sheet sheet,int i,int j,int colcur,int colcnt) throws IOException {//普通input text
	   //获取卷标名称
	   String  labelStr=mapColLable.get(i+","+colcnt);
	   Row row = sheet.getRow(i);
	   Cell cell = row.getCell(j);
	   String  field=getCellContent(cell); //字段名称
	   if (colcur==1) {
		   bw.newLine();
		   bw.write(tab3+"<el-row>");
		   bw.newLine();
		   bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">");   //:span="8"
		   bw.newLine();
		   bw.write(tab5+"<el-form-item  label=\""+labelStr+"\"  prop=\""+field+"\">");
		   bw.newLine();
		   bw.write(tab6+"<el-input type=\"textarea\"  v-model=\"editForm."+field+"\"  placeholder=\"请输入"+labelStr+"\"   :autosize=\"{ minRows: 6, maxRows: 10}\" :maxlength=\"400\"></el-input>");
		   bw.newLine();
		   bw.write(tab5+"</el-form-item>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
	   }else {
		   bw.newLine();
		   bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">");   //:span="8"
		   bw.newLine();
		   bw.write(tab5+"<el-form-item  label=\""+labelStr+"\"  prop=\""+field+"\">");
		   bw.newLine();
		   bw.write(tab6+"<el-input type=\"textarea\"  v-model=\"editForm."+field+"\"  placeholder=\"请输入"+labelStr+"\"   :autosize=\"{ minRows: 6, maxRows: 10}\" :maxlength=\"400\"></el-input>");
		   bw.newLine();
		   bw.write(tab5+"</el-form-item>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
	   }
	   
	   if  (colcur==mapColI.get(i) || mapColI.get(i)==null) {
		   bw.newLine();
		   bw.write(tab3+"</el-row>");
	   }
	   
   }
   
   /***
    * 处理创建数字控制
    * @param bw  sheet  i  j colcur colcnt
    * @throws IOException
    */
   public static void  createInputNum(BufferedWriter bw,Sheet sheet,int i,int j,int colcur,int colcnt) throws IOException {//普通input text
	   //获取卷标名称
	   String  labelStr=mapColLable.get(i+","+colcnt);
	   Row row = sheet.getRow(i);
	   Cell cell = row.getCell(j);
	   String  field=getCellContent(cell); //字段名称
	   if (colcur==1) {
		   bw.newLine();
		   bw.write(tab3+"<el-row>");
		   bw.newLine();
		   bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">");   //:span="8"
		   bw.newLine();
		   bw.write(tab5+"<el-form-item  label=\""+labelStr+"\"  prop=\""+field+"\">");
		   bw.newLine();
		   bw.write(tab6+"<el-input-number  v-model=\"editForm."+field+"\"  placeholder=\"请输入"+labelStr+"\"   :min=\"1\" :max=\"10\"></el-input>");
		   bw.newLine();
		   bw.write(tab5+"</el-form-item>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
		   //<el-input-number v-model="num1" @change="handleChange" :min="1" :max="10" label="描述文字"></el-input-number>

	   }else {
		   bw.newLine();
		   bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">");   //:span="8"
		   bw.newLine();
		   bw.write(tab5+"<el-form-item  label=\""+labelStr+"\"  prop=\""+field+"\">");
		   bw.newLine();
		   bw.write(tab6+"<el-input-number  v-model=\"editForm."+field+"\"  placeholder=\"请输入"+labelStr+"\"   :min=\"1\" :max=\"10\"></el-input>");
		   bw.newLine();
		   bw.write(tab5+"</el-form-item>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
	   }
	   
	   if  (colcur==mapColI.get(i) || mapColI.get(i)==null) {
		   bw.newLine();
		   bw.write(tab3+"</el-row>");
	   }
	   
   }
   
   /***
    * 处理创建单选按钮
    * @param bw  sheet  i  j colcur colcnt
    * @throws IOException
    */
   public static void  createRadio(BufferedWriter bw,Sheet sheet,int i,int j,int colcur,int colcnt) throws IOException {//普通input text
	   //获取卷标名称
	   String  labelStr=mapColLable.get(i+","+colcnt);
	   Row row = sheet.getRow(i);
	   Cell cell = row.getCell(j);
	   String  cellStr=getCellContent(cell);
	   //------------------------------------------行拆分
	   String[] splits=cellStr.split("/r/n") ;
	   //------------------------------------------
	   String  field=splits[0]; //字段名称
	   	   
	   if (colcur==1) {
		   bw.newLine();
		   bw.write(tab3+"<el-row>");
		   bw.newLine();
		   bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">");   //:span="8"
		   bw.newLine();
		   bw.write(tab5+"<el-form-item  label=\""+labelStr+"\"  prop=\""+field+"\">");
		   bw.newLine();
		   bw.write(tab6+"<el-radio-group  v-model=\"editForm."+field+"\"  placeholder=\"请输入"+labelStr+"\"");
		   if  (cellStr.indexOf("#") >=0) {//单选按钮
			   for  (int k=1;k<splits.length;k++) {
				     String[]  strs=splits[k].split("#");
				     bw.newLine();
				     bw.write(tab7+"<el-radio"+ strs[0]+">"+strs[1]+"</el-radio>");
			   }
		   }else { //radio in 方式
			   bw.write(tab7+"<el-radio ");
			   for  (int k=1;k<splits.length;k++) {
				     bw.newLine();
				     bw.write(tab8+ splits[k]);
			   }
			   bw.newLine();
			   bw.write(tab7+">");
			   bw.newLine();
			   bw.write(tab7+"</el-radio>");
		   }
		   bw.newLine();
		   bw.write(tab6+"</el-radio-group>");
		   //---------------------------------
		   bw.newLine();
		   bw.write(tab5+"</el-form-item>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
	   }else {
		   bw.newLine();
		   bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">");   //:span="8"
		   bw.newLine();
		   bw.write(tab5+"<el-form-item  label=\""+labelStr+"\"  prop=\""+field+"\">");
		   bw.newLine();
		   bw.write(tab6+"<el-radio-group  v-model=\"editForm."+field+"\"  placeholder=\"请输入"+labelStr+"\"");
		   if  (cellStr.indexOf("#") >=0) {//单选按钮
			   for  (int k=1;k<splits.length;k++) {
				     String[]  strs=splits[k].split("#");
				     bw.newLine();
				     bw.write(tab7+"<el-radio  "+ strs[0]+">"+strs[1]+"</el-radio>");
			   }
		   }else { //radio in 方式
			   bw.write(tab7+"<el-radio ");
			   for  (int k=1;k<splits.length;k++) {
				     bw.newLine();
				     bw.write(tab8+ splits[k]);
			   }
			   bw.newLine();
			   bw.write(tab9+">");
			   bw.newLine();
			   bw.write(tab10+"</el-radio>");
		   }
		   bw.newLine();
		   bw.write(tab6+"</el-radio-group>");
		   //---------------------------------
		   bw.newLine();
		   bw.write(tab5+"</el-form-item>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
	   }
	   
	   if  (colcur==mapColI.get(i) || mapColI.get(i)==null) {
		   bw.newLine();
		   bw.write(tab3+"</el-row>");
	   }
	   
   }
   
   /***
    * 处理创建checkbox
    * @param bw  sheet  i  j colcur colcnt
    * @throws IOException
    */
   public static void  createCheckbox(BufferedWriter bw,Sheet sheet,int i,int j,int colcur,int colcnt) throws IOException {//普通input text
	   //获取卷标名称
	   String  labelStr=mapColLable.get(i+","+colcnt);
	   Row row = sheet.getRow(i);
	   Cell cell = row.getCell(j);
	   String  cellStr=getCellContent(cell);
	   //------------------------------------------行拆分
	   String[] splits=cellStr.split("/r/n") ;
	   //------------------------------------------
	   String  field=splits[0]; //字段名称
	   	   
	   if (colcur==1) {
		   bw.newLine();
		   bw.write(tab3+"<el-row>");
		   bw.newLine();
		   bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">");   //:span="8"
		   bw.newLine();
		   bw.write(tab5+"<el-form-item  label=\""+labelStr+"\"  prop=\""+field+"\">");
		   bw.newLine();
		   bw.write(tab6+"<el-checkbox-group  v-model=\"editForm."+field+"\"  placeholder=\"请输入"+labelStr+"\"");
		   if  (cellStr.indexOf("#") >=0) {//单选按钮
			   for  (int k=1;k<splits.length;k++) {
				     String[]  strs=splits[k].split("#");
				     bw.newLine();
				     bw.write(tab7+"<el-checkbox   "+ strs[0]+">"+strs[1]+"</el-checkbox>");
			   }
		   }else { //radio in 方式
			   bw.write(tab7+"<el-checkbox ");
			   for  (int k=1;k<splits.length;k++) {
				     bw.newLine();
				     bw.write(tab8+ splits[k]);
			   }
			   bw.newLine();
			   bw.write(tab7+">");
			   bw.newLine();
			   bw.write(tab7+"</el-checkbox>");
		   }
		   bw.newLine();
		   bw.write(tab6+"</el-checkbox-group>");
		   //---------------------------------
		   bw.newLine();
		   bw.write(tab5+"</el-form-item>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
	   }else {
		   bw.newLine();
		   bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">");   //:span="8"
		   bw.newLine();
		   bw.write(tab5+"<el-form-item  label=\""+labelStr+"\"  prop=\""+field+"\">");
		   bw.newLine();
		   bw.write(tab6+"<el-checkbox-group  v-model=\"editForm."+field+"\"  placeholder=\"请输入"+labelStr+"\"");
		   if  (cellStr.indexOf("#") >=0) {//单选按钮
			   for  (int k=1;k<splits.length;k++) {
				     String[]  strs=splits[k].split("#");
				     bw.newLine();
				     bw.write(tab7+"<el-checkbox   "+ strs[0]+">"+strs[1]+"</el-checkbox>");
			   }
		   }else { //radio in 方式
			   bw.write(tab7+"<el-checkbox ");
			   for  (int k=1;k<splits.length;k++) {
				     bw.newLine();
				     bw.write(tab8+ splits[k]);
			   }
			   bw.newLine();
			   bw.write(tab9+">");
			   bw.newLine();
			   bw.write(tab10+"</el-checkbox>");
		   }
		   bw.newLine();
		   bw.write(tab6+"</el-checkbox-group>");
		   //---------------------------------
		   bw.newLine();
		   bw.write(tab5+"</el-form-item>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
	   }
	   
	   if  (colcur==mapColI.get(i) || mapColI.get(i)==null) {
		   bw.newLine();
		   bw.write(tab3+"</el-row>");
	   }
	   
   }
   
   /***
    * 处理创建select
    * @param bw  sheet  i  j colcur colcnt
    * @throws IOException
    */
   public static void  createSelect(BufferedWriter bw,Sheet sheet,int i,int j,int colcur,int colcnt) throws IOException {//普通input text
	   //获取卷标名称
	   String  labelStr=mapColLable.get(i+","+colcnt);
	   Row row = sheet.getRow(i);
	   Cell cell = row.getCell(j);
	   String  cellStr=getCellContent(cell);
	   //------------------------------------------行拆分
	   String[] splits=cellStr.split("/r/n") ;
	   //------------------------------------------
	   String  field=splits[0]; //字段名称
	   	   
	   if (colcur==1) {
		   bw.newLine();
		   bw.write(tab3+"<el-row>");
		   bw.newLine();
		   bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">");   //:span="8"
		   bw.newLine();
		   bw.write(tab5+"<el-form-item  label=\""+labelStr+"\"  prop=\""+field+"\">");
		   bw.newLine();
		   bw.write(tab6+"<el-select  v-model=\"editForm."+field+"\"  placeholder=\"请选择"+labelStr+"\"");
		   if  (cellStr.indexOf("#") >=0) {//单选按钮
			   for  (int k=1;k<splits.length;k++) {
				     bw.newLine();
				     bw.write(tab7+"<el-option  "+ splits[k]+"></el-option>");
			   }
		   }else { //Select in 方式
			   bw.write(tab7+"<el-option ");
			   for  (int k=1;k<splits.length;k++) {
				     bw.newLine();
				     bw.write(tab8+ splits[k]);
			   }
			   bw.newLine();
			   bw.write(tab7+">");
			   bw.newLine();
			   bw.write(tab7+"</el-option>");
		   }
		   bw.newLine();
		   bw.write(tab6+"</el-select>");
		   //---------------------------------
		   bw.newLine();
		   bw.write(tab5+"</el-form-item>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
	   }else {
		   bw.newLine();
		   bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">");   //:span="8"
		   bw.newLine();
		   bw.write(tab5+"<el-form-item  label=\""+labelStr+"\"  prop=\""+field+"\">");
		   bw.newLine();
		   bw.write(tab6+"<el-select  v-model=\"editForm."+field+"\"  placeholder=\"请输入"+labelStr+"\"");
		   if  (cellStr.indexOf("#") >=0) {//单选按钮
			   for  (int k=1;k<splits.length;k++) {
				     bw.newLine();
				     bw.write(tab7+"<el-option   "+ splits[k]+"></el-option>");
			   }
		   }else { //radio in 方式
			   bw.write(tab7+"<el-option ");
			   for  (int k=1;k<splits.length;k++) {
				     bw.newLine();
				     bw.write(tab8+ splits[k]);
			   }
			   bw.newLine();
			   bw.write(tab9+">");
			   bw.newLine();
			   bw.write(tab10+"</el-option>");
		   }
		   bw.newLine();
		   bw.write(tab6+"</el-select>");
		   //---------------------------------
		   bw.newLine();
		   bw.write(tab5+"</el-form-item>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
	   }
	   
	   if  (colcur==mapColI.get(i) || mapColI.get(i)==null) {
		   bw.newLine();
		   bw.write(tab3+"</el-row>");
	   }
	   
   }
   
   /***
    * 处理创建switch
    * @param bw  sheet  i  j colcur colcnt
    * @throws IOException
    */
   public static void  createSwitch(BufferedWriter bw,Sheet sheet,int i,int j,int colcur,int colcnt) throws IOException {//普通input text
	   //获取卷标名称
	   String  labelStr=mapColLable.get(i+","+colcnt);
	   Row row = sheet.getRow(i);
	   Cell cell = row.getCell(j);
	   String  cellStr=getCellContent(cell);
	   //------------------------------------------行拆分
	   String[] splits=cellStr.split("/r/n") ;
	   //------------------------------------------
	   String  field=splits[0]; //字段名称	   
	   if (colcur==1) {
		   bw.newLine();
		   bw.write(tab3+"<el-row>");
		   bw.newLine();
		   bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">");   //:span="8"
		   bw.newLine();
		   bw.write(tab5+"<el-form-item  label=\""+labelStr+"\"  prop=\""+field+"\">");

		   bw.write(tab6+"<el-switch ");
		   for  (int k=1;k<splits.length;k++) {
			     bw.newLine();
			     bw.write(tab7+ splits[k]);
		   }
		   bw.newLine();
		   bw.write(tab8+">");
		   bw.newLine();
		   bw.write(tab9+"</el-switch>");
		   //---------------------------------
		   bw.newLine();
		   bw.write(tab5+"</el-form-item>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
	   }else {
		   bw.newLine();
		   bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">");   //:span="8"
		   bw.newLine();
		   bw.write(tab5+"<el-form-item  label=\""+labelStr+"\"  prop=\""+field+"\">");
		   bw.write(tab6+"<el-switch ");
		   for  (int k=1;k<splits.length;k++) {
			     bw.newLine();
			     bw.write(tab7+ splits[k]);
		   }
		   bw.newLine();
		   bw.write(tab8+">");
		   bw.newLine();
		   bw.write(tab9+"</el-switch>");
		   //---------------------------------
		   bw.newLine();
		   bw.write(tab5+"</el-form-item>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
	   }
	   
	   if  (colcur==mapColI.get(i) || mapColI.get(i)==null) {
		   bw.newLine();
		   bw.write(tab3+"</el-row>");
	   }
	   
   }
   
   /***
    * 处理创建date
    * @param bw  sheet  i  j colcur colcnt
    * @throws IOException
    */
   public static void  createDate(BufferedWriter bw,Sheet sheet,int i,int j,int colcur,int colcnt) throws IOException {//普通input text
	   //获取卷标名称
	   String  labelStr=mapColLable.get(i+","+colcnt);
	   Row row = sheet.getRow(i);
	   Cell cell = row.getCell(j);
	   String  field=getCellContent(cell); //字段名称
	   if (colcur==1) {
		   bw.newLine();
		   bw.write(tab3+"<el-row>");
		   bw.newLine();
		   bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">");   //:span="8"
		   bw.newLine();
		   if  (labelStr!=null && !"".equals(labelStr)) {
		       bw.write(tab5+"<el-form-item  label=\""+labelStr+"\"  prop=\""+field+"\">");
		   }else {
			   bw.write(tab5+"<el-form-item   prop=\""+field+"\">");
		   }
		   bw.newLine();
		   bw.write(tab6+"<el-date-picker type=\"date\" v-model=\"editForm."+field+"\"  placeholder=\"请选择日期"+"\"   style=\"width: 100%;\"></el-date-picker>");
		   bw.newLine();
		   bw.write(tab5+"</el-form-item>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
	   }else {
		   bw.newLine();
		   bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">");   //:span="8"
		   bw.newLine();
		   if  (labelStr!=null && !"".equals(labelStr)) {
		       bw.write(tab5+"<el-form-item  label=\""+labelStr+"\"  prop=\""+field+"\">");
		   }else {
			   bw.write(tab5+"<el-form-item   prop=\""+field+"\">");
		   }
		   bw.newLine();
		   bw.write(tab6+"<el-date-picker type=\"date\" v-model=\"editForm."+field+"\"  placeholder=\"请选择日期"+"\"   style=\"width: 100%;\"></el-date-picker>");
		   bw.newLine();
		   bw.write(tab5+"</el-form-item>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
	   }
	   
	   if  (colcur==mapColI.get(i) || mapColI.get(i)==null) {
		   bw.newLine();
		   bw.write(tab3+"</el-row>");
	   }
	   
   }
   
   /***
    * 处理创建Time
    * @param bw  sheet  i  j colcur colcnt
    * @throws IOException
    */
   public static void  createTime(BufferedWriter bw,Sheet sheet,int i,int j,int colcur,int colcnt) throws IOException {//
	   //获取卷标名称
	   String  labelStr=mapColLable.get(i+","+colcnt);
	   Row row = sheet.getRow(i);
	   Cell cell = row.getCell(j);
	   String  field=getCellContent(cell); //字段名称
	   if (colcur==1) {
		   bw.newLine();
		   bw.write(tab3+"<el-row>");
		   bw.newLine();
		   bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">");   //:span="8"
		   bw.newLine();
		   if  (labelStr!=null && !"".equals(labelStr)) {
		        bw.write(tab5+"<el-form-item  label=\""+labelStr+"\"  prop=\""+field+"\">");
		   }else {
			   bw.write(tab5+"<el-form-item  prop=\""+field+"\">");
		   }
		   bw.newLine();
		   bw.write(tab6+"<el-time-picker type=\"date\" v-model=\"editForm."+field+"\"  placeholder=\"请选择时间"+"\"   style=\"width: 100%;\"></el-time-picker>");
		   bw.newLine();
		   bw.write(tab5+"</el-form-item>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
	   }else {
		   bw.newLine();
		   bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">");   //:span="8"
		   bw.newLine();
		   if  (labelStr!=null && !"".equals(labelStr)) {
		        bw.write(tab5+"<el-form-item  label=\""+labelStr+"\"  prop=\""+field+"\">");
		   }else {
			   bw.write(tab5+"<el-form-item  prop=\""+field+"\">");
		   }
		   bw.newLine();
		   bw.write(tab6+"<fixed-time type=\"date\" v-model=\"editForm."+field+"\"  placeholder=\"请选择时间"+"\"   style=\"width: 100%;\"></fixed-time>");
		   bw.newLine();
		   bw.write(tab5+"</el-form-item>");
		   bw.newLine();
		   bw.write(tab4+"</el-col>");
	   }
	   
	   if  (colcur==mapColI.get(i) || mapColI.get(i)==null) {
		   bw.newLine();
		   bw.write(tab3+"</el-row>");
	   }
	   
   }
   
   /***
    * 处理创建Span  
    * @param bw  sheet  i  j colcur colcnt
    * @throws IOException
    */
   public static void  createSpan(BufferedWriter bw,Sheet sheet,int i,int j,int colcur,int colcnt) throws IOException {//
	   //获取卷标名称
	   //String  labelStr=mapColLable.get(i+","+colcnt);
	   Row row = sheet.getRow(i);
	   Cell cell = row.getCell(j);
	   String  field=getCellContent(cell); //
	   if (colcur==1) {
		   bw.newLine();
		   bw.write(tab3+"<el-row>");
		   bw.newLine();
		   if (mapColIcur.get(i+","+colcnt)!=null) {
		      bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">"+ "</el-col>");   //:span="8"
		   }else {
			  bw.write(tab4+"<el-col :span=\"24"+"\">"+ "</el-col>");   //:span="8"
		   }
	   }else {
		   bw.newLine();
		   if (mapColIcur.get(i+","+colcnt)!=null) {
		      bw.write(tab4+"<el-col :span=\""+mapColIcur.get(i+","+colcnt)+"\">"+ "</el-col>");   //:span="8"
		   }else {
			  bw.write(tab4+"<el-col :span=\"24"+"\">"+ "</el-col>");   //:span="8"
		   }
	   }
	   if  (colcur==mapColI.get(i) || mapColI.get(i)==null) {
		   bw.newLine();
		   bw.write(tab3+"</el-row>");
	   }
	   
   }
   
   
   /***
    * 预处理数据项
    * @param sheet
    * @param rows
    */
   public  static void setPreCol(Sheet sheet,int rows) {
	   //----------------------------------------------
       for (int i = 5; i < rows; i++) {
    	   Row row = sheet.getRow(i);
    	   if (row != null) {
               // 获取到Excel文件中的所有的列
               int cells = row.getPhysicalNumberOfCells();
               // 遍历列
                int  i_lable=0;
                for (int j = 0; j < cells; j++) {
                   // 获取到列的值
                   Cell cell = row.getCell(j);
                   if (cell != null &&  getCellContent(cell)!=null && !"".equals(getCellContent(cell)) ) {
                	  if (cell.getCellStyle().getFillForegroundColorColor()==null) {//非标签类 则要占col
                		  if  (mapColI.get(i)==null) {
                			   mapColI.put(i, 1);
                		  }else {
                			  mapColI.put(i, mapColI.get(i)+1);//列数+1
                		  }
                	  } else {//登记卷标
                	    i_lable=i_lable + 1;
                	    mapColLable.put(i+","+i_lable,getCellContent(cell)) ;
                      }
                   }
                	  
                }
                
                //计算出每个格子 所占列宽
                if  (mapColI.get(i)!=null) {
                	 //Math.floor(double a)
                	int  icol=(new  Double(Math.floor(24.0/mapColI.get(i)))).intValue() ;       //  Math.floor(24.0/mapColI.get(i))).intValue();
                	//(new   Double(d)).intValue();
                	for  (int k=0;k< mapColI.get(i) - 1;k++ ) {
                		  mapColIcur.put(i+","+(k+1), icol);
                	}
                	mapColIcur.put(i+","+(mapColI.get(i)), 24 - (mapColI.get(i) - 1)*icol); //行列最后一个
                }
                //-----------------------------------------------
    	   }
       }
       System.out.println(mapColI);
       System.out.println(mapColIcur);
       System.out.println(mapColLable);
       //----------------------------------------------
	   
   } 
}