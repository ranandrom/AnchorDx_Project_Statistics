package ProjectStatistics;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * Description: <br/>
 * Function: This class is project statistics class. <br/>
 * File Name: AnchorDx_Project_Statistics.java <br/>
 * Date: 2017-04-25
 * 
 * @author Luzhirong ramandrom@139.com
 * @version V1.0.0
 */

public class ProjectStatistics
{
	/**
	 * projectStatistics程序的入口.
	 */
	public static void projectStatisticsMain(String OutputPath, String system)
	{
		Calendar now_star = Calendar.getInstance();		
		SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd");//定义日期格式
		String ToDay = formatter.format(now_star.getTime());//获取当天日期

		myMkdir(OutputPath);
		String OutputFileName = null;
		if (system.equals("linux")) {
			OutputFileName = OutputPath + "/" + "项目统计表_" + ToDay + ".xlsx";
		} else {
			OutputFileName = OutputPath + "\\" + "项目统计表_" + ToDay + ".xlsx";
		}
		File OutPutfile = new File(OutputFileName);
		createXlsx(OutPutfile); //创建项目统计表		
		ArrayList<String> filelist = new ArrayList<String>();
		searchStatisticsFile(new File(OutputPath), filelist, system); // 获取需要统计的文件
		
		String data = null;
		ArrayList<String> datalist = new ArrayList<String>();
		for (int i = 0; i < filelist.size(); i++) {
			String filepath[] = filelist.get(i).split("\t");
			data = statisticsFile(filepath[0], filepath[1]);
			datalist.add(data);
		}
		writeDataToStatistics(OutputFileName, datalist); // 写数据到统计表
	}
	
	/**
	 * 创建目录的方法
	 * @param dir_name
	 */
	public static void myMkdir(String dir_name)
	{
		File file = new File(dir_name);
		//如果文件不存在，则创建
		if (!file.exists() && !file.isDirectory()) {
			file.mkdirs();
		}
	}
	
	/**
	 * 创建一个统计表的方法。
	 * @param file
	 */
	@SuppressWarnings("deprecation")
	public static void createXlsx(File file)
	{
		try {
			XSSFWorkbook Workbook = new XSSFWorkbook();
			// 创建Excel的工作sheet,对应到一个excel文档的tab  
			XSSFSheet sheet = Workbook.createSheet("所有项目统计");
			// 在索引0的位置创建行（最顶端的行）
			XSSFRow row0 = sheet.createRow((short) 0);

			String head_row0 = "Porject name" + "\t" +"Plasma already analyzed" + "\t" + "Plasma not analyzed" + "\t" + "Plasma failed" + "\t"
								+ "Tissue already analyzed" + "\t"+ "Tissue not analyzed" + "\t" + "Tissue failed" + "\t"
								+ "WBC already analyzed" + "\t" + "WBC not analyzed" + "\t" + "WBC failed" + "\t" + "Porject sum";
			
			// 1、创建字体，设置字体为粗体，背景绿色：
			XSSFFont font0 = Workbook.createFont();
			font0.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
			font0.setFontHeightInPoints((short) 10);
			font0.setFontName("Palatino Linotype");
			// 2、创建格式
			XSSFCellStyle cellStyle0 = Workbook.createCellStyle();
			cellStyle0.setFont(font0);
			cellStyle0.setBorderTop(XSSFCellStyle.BORDER_THIN);
			cellStyle0.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle0.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle0.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle0.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
			cellStyle0.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			
			//1、创建字体，设置字体为粗体，背景蓝色：
			XSSFFont font1 = Workbook.createFont();
			font1.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
			font1.setFontHeightInPoints((short)10);
			font1.setFontName("Palatino Linotype");
			//2、创建格式
			XSSFCellStyle cellStyle1= Workbook.createCellStyle();
			cellStyle1.setFont(font1);
			cellStyle1.setBorderTop(XSSFCellStyle.BORDER_THIN);
			cellStyle1.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle1.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle1.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle1.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
			cellStyle1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			
			// 1、创建字体，设置其为粗体、紅色，背景黄色：
			XSSFFont font2 = Workbook.createFont();
			font2.setColor(HSSFFont.COLOR_RED);
			font2.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
			font2.setFontHeightInPoints((short) 10);
			font2.setFontName("Palatino Linotype");
			// 2、创建格式
			XSSFCellStyle cellStyle2 = Workbook.createCellStyle();
			cellStyle2.setFont(font2);
			cellStyle2.setBorderTop(XSSFCellStyle.BORDER_THIN);
			cellStyle2.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle2.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle2.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle2.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
			cellStyle2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			
			// 1、创建字体，设置其为粗体，背景红色：
			XSSFFont font3 = Workbook.createFont();
			font3.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
			font3.setFontHeightInPoints((short) 10);
			font3.setFontName("Palatino Linotype");
			// 2、创建格式
			XSSFCellStyle cellStyle3 = Workbook.createCellStyle();
			cellStyle3.setFont(font3);
			cellStyle3.setBorderTop(XSSFCellStyle.BORDER_THIN);
			cellStyle3.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle3.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle3.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle3.setFillForegroundColor(HSSFColor.RED.index);
			cellStyle3.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			
			// 1、创建字体，设置其为粗体，背景黄色：
			XSSFFont font4 = Workbook.createFont();
			font4.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
			font4.setFontHeightInPoints((short) 10);
			font4.setFontName("Palatino Linotype");
			// 2、创建格式
			XSSFCellStyle cellStyle4 = Workbook.createCellStyle();
			cellStyle4.setFont(font4);
			cellStyle4.setBorderTop(XSSFCellStyle.BORDER_THIN);
			cellStyle4.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle4.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle4.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle4.setFillForegroundColor(HSSFColor.LIGHT_ORANGE.index);
			cellStyle4.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);


			String str_head_row0[] = head_row0.split("\t");
			
			// 往单元格中输入一些内容
			for (int i = 0; i < str_head_row0.length; i++) {
				// 在索引0的位置创建单元格（左上端）
				XSSFCell cell = row0.createCell(i);
				if (i == 0) { // 表格的已分析单元格为粗体字体，背景绿色
					cell.setCellValue(str_head_row0[i]);
					cell.setCellStyle(cellStyle0);
				} else if (str_head_row0[i].contains("already")) { // 表格未分析单元格为粗体、紅色字体，背景黄色
					cell.setCellStyle(cellStyle1);
					cell.setCellValue(str_head_row0[i]);
				} else if (str_head_row0[i].contains("not")) {
					cell.setCellStyle(cellStyle2);
					cell.setCellValue(str_head_row0[i]);
				} else if (str_head_row0[i].contains("sum")){
					cell.setCellStyle(cellStyle4);
					cell.setCellValue(str_head_row0[i]);
				} else {
					cell.setCellStyle(cellStyle3);
					cell.setCellValue(str_head_row0[i]);
				}
			}

			// 新建一输出文件流
			FileOutputStream fOut = new FileOutputStream(file);
			// 把相应的Excel工作簿存盘
			Workbook.write(fOut);
			fOut.flush();
			// 操作结束，关闭文件
			fOut.close();
			Workbook.close();
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	/**
	 * 获取最新文件所在目录
	 * @param Path
	 * @return
	 */
	public static String getNewestFileDir(String Path)
	{
		File file = new File(Path);
		int daynum = 0;
		for (File dir : file.listFiles()){
			if (dir.isDirectory()) { //如果是目录
				String dir_name = dir.getName(); //目录名
				if (daynum < Integer.valueOf(dir_name)) {
					daynum = Integer.valueOf(dir_name);
				} else {
					continue;
				}
			} else {
				continue;
			}
		}
		return String.valueOf(daynum);
	}
	
	/**
	 * 查找需要统计的项目文件
	 * @param des_file
	 * @param list
	 */
	public static void searchStatisticsFile(File des_file, ArrayList<String> filelist, String system)
    {	
		try{ 
			for (File pathname : des_file.listFiles())
			{
				if (pathname.isFile()) //如果是文件，则判断是否需要记录
				{
					//获取文件的绝对路径
					String Folder = pathname.getParent();					
					//获取文件名（basename）
					String FileName = pathname.getName();
					
					if (!(FileName.startsWith("~$")) && !(FileName.startsWith("."))) {
						String FileName_str[] = FileName.split("_");
						if (FileName_str[1].equals("项目实验生信汇总表")) {
							String str = null;
							if (system.equals("linux")) {
								str = Folder + "/" + FileName + "\t" + FileName_str[0];
							} else {
								str = Folder + "\\" + FileName + "\t" + FileName_str[0];
							}
							
							filelist.add(str);
						}
					}
				} else {
					//如果是目录，则递归
					searchStatisticsFile(pathname, filelist, system);
				}
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }
	
	/**
	 * 统计并生成统计表文件
	 * @param filelist
	 */
	@SuppressWarnings({ "deprecation", "resource" })
	public static String statisticsFile(String filepath, String PorjectName)
	{
		String data = null;
		try {
			File file = new File(filepath);
			InputStream is = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(is);
			//XSSFSheet sheet = wb.getSheetAt(2);	// 获取第三个工作薄
			XSSFSheet sheet = null;
			int Sheet_Num = wb.getNumberOfSheets(); // 获取工作薄个数
			int Plasma_already_analyzed_all = 0; // 所有的血浆已分析
			int Plasma_not_analyzed_all = 0; // 所有的血浆未分析
			int Plasma_failed_all = 0; // 所有的血浆失败
			int Tissue_already_analyzed_all = 0; // 所有的组织已分析
			int Tissue_not_analyzed_all = 0; // 所有的组织未分析
			int Tissue_failed_all = 0; // 所有的组织失败
			int WBC_already_analyzed_all = 0; // 所有的白细胞已分析
			int WBC_not_analyzed_all = 0; // 所有的白细胞未分析
			int WBC_failed_all = 0; // 所有的白细胞失败
			int porject_sum = 0; // 项目总和
			
			for(int numSheet = 0; numSheet < Sheet_Num; numSheet++ ){
				int Plasma_already_analyzed = 0; // 本页的血浆已分析
				int Plasma_not_analyzed = 0; // 本页的血浆未分析
				int Plasma_failed = 0; // 本页的血浆失败
				int Tissue_already_analyzed = 0; // 本页的组织已分析
				int Tissue_not_analyzed = 0; // 本页的组织未分析
				int Tissue_failed = 0; // 本页的组织失败
				int WBC_already_analyzed = 0; // 本页的白细胞已分析
				int WBC_not_analyzed = 0; // 本页的白细胞未分析
				int WBC_failed = 0; // 本页的白细胞失败
				
				sheet = wb.getSheetAt(numSheet);	// 获取工作薄
				//String Sheet_Name = sheet.getSheetName(); // 获取当前工作薄名字
				
				XSSFRow xssfrow0 = sheet.getRow(0); //获取第一行
				int ExperimentColumn = 0;
				int BioinfoColumn = 0;
				int log = 0;
				int failed = 0; // 失败次数
				int Plasma_log = 0; // 血浆标志
				int Tissue_log = 0; // 组织标志
				int WBC_log = 0; // 白细胞标志
				// 获取当前工作薄 "Sample ID"列后面的那列(需要统计的列)。
				for (int j = xssfrow0.getFirstCellNum(); j < xssfrow0.getLastCellNum(); j++) {
					XSSFCell xssfcell = xssfrow0.getCell(j);
					int xcellty = xssfcell.getCellType();
					String head = null;
					if (xcellty == 0) {
						if (HSSFDateUtil.isCellDateFormatted(xssfcell)) { // 判断是否为日期格式
							Date date = xssfcell.getDateCellValue();
							SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
							head = dateFormat.format(date);
						} else {
							xssfcell.setCellType(HSSFCell.CELL_TYPE_NUMERIC); // 设置单元格为数值类型
							HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
							String cellFormatted = dataFormatter.formatCellValue(xssfcell); // 按单元格数据格式类型读取数据
							head = cellFormatted;
						}
					} else {
						xssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
						head = xssfcell.getStringCellValue().trim();
					}
					if (head.equals("Sample ID")) {
						if (log == 1) {
							BioinfoColumn = j+1;
							break;
						} else {
							ExperimentColumn = j+1;
						}
						log++;
					}
				}
				
				// 确定该工作页是组织、血浆、白细胞中的哪一个的数据
				for (int i = sheet.getFirstRowNum()+1; i <= sheet.getLastRowNum(); i++) {
					XSSFRow xssfrow = sheet.getRow(i); // 获取每一行
					XSSFCell xssfcell_Experiment = xssfrow.getCell(ExperimentColumn); // 获取实验数据“Pre lib name”所在列的单元格数据
					int xcellty_Experiment = xssfcell_Experiment.getCellType();
					String Experiment_data = null;
					if (xcellty_Experiment == 0) {
						if (HSSFDateUtil.isCellDateFormatted(xssfcell_Experiment)) { // 判断是否为日期格式
							Date date = xssfcell_Experiment.getDateCellValue();
							SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
							Experiment_data = dateFormat.format(date);
						} else {
							xssfcell_Experiment.setCellType(HSSFCell.CELL_TYPE_NUMERIC); // 设置单元格为数值类型
							HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
							String cellFormatted = dataFormatter.formatCellValue(xssfcell_Experiment); // 按单元格数据格式类型读取数据
							Experiment_data = cellFormatted;
						}
					} else {
						xssfcell_Experiment.setCellType(HSSFCell.CELL_TYPE_STRING);
						Experiment_data = xssfcell_Experiment.getStringCellValue().trim();
					}
					
					XSSFCell xssfcell_Bioinfo = xssfrow.getCell(BioinfoColumn); // 获取生信数据“Pre lib name”所在列的单元格数据
					int xcellty_Bioinfo = -1;
					String Bioinfo_data = null;
					try {
						xcellty_Bioinfo = xssfcell_Bioinfo.getCellType();
					} catch (Exception e) {
						if (Experiment_data == null || Experiment_data.equals("")) {
							continue;
						} else {
							String Pre_lib_name[] = Experiment_data.split("-");
							if (Pre_lib_name.length > 3) {
								if (Pre_lib_name[Pre_lib_name.length-1].startsWith("PS")) {
									Plasma_log++;
									Plasma_not_analyzed++;
									continue;
								} else if (Pre_lib_name[Pre_lib_name.length-1].startsWith("F")) {
									Tissue_log++;
									Tissue_not_analyzed++;
									continue;
								} else if (Pre_lib_name[Pre_lib_name.length-1].startsWith("BC")){
									WBC_log++;
									WBC_not_analyzed++;
									continue;
								} else {
									continue;
								}
							} else {
								failed++;
								//System.out.println(Experiment_data + failed);
								continue;
							}
						}
					}
					//System.out.println("xcellty_Bioinfo== " + xcellty_Bioinfo);
					if (xcellty_Bioinfo == 0) {
						if (HSSFDateUtil.isCellDateFormatted(xssfcell_Bioinfo)) { // 判断是否为日期格式
							Date date = xssfcell_Bioinfo.getDateCellValue();
							SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
							Bioinfo_data = dateFormat.format(date);
						} else {
							xssfcell_Bioinfo.setCellType(HSSFCell.CELL_TYPE_NUMERIC); // 设置单元格为数值类型
							HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
							String cellFormatted = dataFormatter.formatCellValue(xssfcell_Bioinfo); // 按单元格数据格式类型读取数据
							Bioinfo_data = cellFormatted;
						}
					} else {
						xssfcell_Bioinfo.setCellType(HSSFCell.CELL_TYPE_STRING);
						Bioinfo_data = xssfcell_Bioinfo.getStringCellValue().trim();
					}
					
					if (Experiment_data == null || Experiment_data.equals("")) {
						continue;
					} else {
						String Pre_lib_name[] = Experiment_data.split("-");
						if (Pre_lib_name.length > 3) {
							if (Pre_lib_name[Pre_lib_name.length-1].startsWith("PS")) {
								//System.out.println(filepath + "::" + Sheet_Name + "血浆");
								//break;
								Plasma_log++;
								if (Bioinfo_data == null || Bioinfo_data.equals("")) {
									Plasma_not_analyzed++;
								} else {
									Plasma_already_analyzed++;
								}
								continue;
							} else if (Pre_lib_name[Pre_lib_name.length-1].startsWith("F")) {
								//System.out.println(filepath + "::" + Sheet_Name + "组织");
								//break;
								Tissue_log++;
								if (Bioinfo_data == null || Bioinfo_data.equals("")) {
									Tissue_not_analyzed++;
								} else {
									Tissue_already_analyzed++;
								}
								continue;
							} else if (Pre_lib_name[Pre_lib_name.length-1].startsWith("BC")){
								//System.out.println(filepath + "::" + Sheet_Name + "白细胞");
								//break;
								WBC_log++;
								if (Bioinfo_data == null || Bioinfo_data.equals("")) {
									WBC_not_analyzed++;
								} else {
									WBC_already_analyzed++;
								}
								continue;
							} else {
								//failed++;
								//System.out.println(Experiment_data);
								continue;
							}
						} else {
							failed++;
							//System.out.println(Experiment_data + "======" + failed);
							continue;
						}
					}
				}
				if (Plasma_log > 0) {
					Plasma_failed = failed;
				} else if (Tissue_log > 0) {
					Tissue_failed = failed;
				}else if (WBC_log > 0) {
					WBC_failed = failed;
				}
				
				//System.out.println(PorjectName + "::" + Sheet_Name);
				//System.out.println("Plasma_already_analyzed = " + Plasma_already_analyzed + "\t" + "Plasma_not_analyzed = " + Plasma_not_analyzed + "\t" + "Plasma_failed = " + Plasma_failed);
				//System.out.println("Tissue_already_analyzed = " + Tissue_already_analyzed + "\t" + "Tissue_not_analyzed = " + Tissue_not_analyzed + "\t" + "Tissue_failed = " + Tissue_failed);
				//System.out.println("WBC_already_analyzed = " + WBC_already_analyzed + "\t" + "WBC_not_analyzed = " + WBC_not_analyzed + "\t" + "WBC_failed = " + WBC_failed);
			
				Plasma_already_analyzed_all += Plasma_already_analyzed; // 所有的血浆已分析
				Plasma_not_analyzed_all += Plasma_not_analyzed; // 所有的血浆未分析
				Plasma_failed_all += Plasma_failed; // 所有的血浆失败
				Tissue_already_analyzed_all += Tissue_already_analyzed; // 所有的组织已分析
				Tissue_not_analyzed_all += Tissue_not_analyzed; // 所有的组织未分析
				Tissue_failed_all += Tissue_failed; // 所有的组织失败
				WBC_already_analyzed_all += WBC_already_analyzed; // 所有的白细胞已分析
				WBC_not_analyzed_all += WBC_not_analyzed; // 所有的白细胞未分析
				WBC_failed_all += WBC_failed; // 所有的白细胞失败
			
			}
			//System.out.println();
			//System.out.println(PorjectName);
			//System.out.println("Plasma_already_analyzed = " + Plasma_already_analyzed_all + "\t" + "Plasma_not_analyzed = " + Plasma_not_analyzed_all + "\t" + "Plasma_failed = " + Plasma_failed_all);
			//System.out.println("Tissue_already_analyzed = " + Tissue_already_analyzed_all + "\t" + "Tissue_not_analyzed = " + Tissue_not_analyzed_all + "\t" + "Tissue_failed = " + Tissue_failed_all);
			//System.out.println("WBC_already_analyzed = " + WBC_already_analyzed_all + "\t" + "WBC_not_analyzed = " + WBC_not_analyzed_all + "\t" + "WBC_failed = " + WBC_failed_all);
			
			porject_sum = Plasma_already_analyzed_all + Plasma_not_analyzed_all + Plasma_failed_all
						+ Tissue_already_analyzed_all + Tissue_not_analyzed_all + Tissue_failed_all
						+ WBC_already_analyzed_all + WBC_not_analyzed_all + WBC_failed_all;
			
			data = PorjectName + "\t"
					+ Plasma_already_analyzed_all + "\t" + Plasma_not_analyzed_all + "\t" + Plasma_failed_all + "\t"
					+ Tissue_already_analyzed_all + "\t" + Tissue_not_analyzed_all + "\t" + Tissue_failed_all + "\t"
					+ WBC_already_analyzed_all + "\t" + WBC_not_analyzed_all + "\t" + WBC_failed_all + "\t"
					+ porject_sum;
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return data;
	}
	
	/**
	 * 写数据到统计表
	 * @param file
	 * @param Data_list
	 */
	public static void writeDataToStatistics(String Outputfile, ArrayList<String> Data_list)
	{	
		try{
			File file = new File(Outputfile);
			FileInputStream is = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(is);
			
			// 1、创建字体，设置其为粗体，背景黄色：
			XSSFFont font4 = workbook.createFont();
			font4.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
			font4.setFontHeightInPoints((short) 10);
			font4.setFontName("Palatino Linotype");
			// 2、创建格式
			XSSFCellStyle cellStyle4 = workbook.createCellStyle();
			cellStyle4.setFont(font4);
			cellStyle4.setBorderTop(XSSFCellStyle.BORDER_THIN);
			cellStyle4.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle4.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle4.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle4.setFillForegroundColor(HSSFColor.LIGHT_ORANGE.index);
			cellStyle4.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			
			//XSSFSheet sheet = workbook.getSheetAt(0);	//获取第1个工作薄
			XSSFSheet sheet = null;
			int Sheet_Num = workbook.getNumberOfSheets();//获取工作薄个数
			//System.out.println(Sheet_Num);
			int sumarr[] = new int[10];
			for(int numSheet = 0; numSheet < Sheet_Num; numSheet++ ){
				sheet = workbook.getSheetAt(numSheet);	//获取工作薄
				//String Sheet_Name = sheet.getSheetName();//获取当前工作薄名字
				//写数据
				for(int j = 0; j < Data_list.size(); j++){
					XSSFRow row = sheet.createRow((short) j+1);
					String str_row[] = Data_list.get(j).split("\t");
					for(int i = 0; i < str_row.length; i++ ){
						// 在索引0的位置创建单元格（左上端）
						XSSFCell cell = row.createCell(i);
						if (i == 0) {
							cell.setCellValue(str_row[i]);
						} else {
							cell.setCellValue(Integer.parseInt(str_row[i]));
							sumarr[i-1] += Integer.parseInt(str_row[i]);
						}
					}
				}
				XSSFRow lastrow = sheet.createRow((short) Data_list.size()+1);
				for(int j = 0; j <= sumarr.length; j++){
					XSSFCell cell = lastrow.createCell(j);
					if (j == 0) {
						cell.setCellStyle(cellStyle4);
						cell.setCellValue("All porject sum");
					} else {
						cell.setCellStyle(cellStyle4);
						cell.setCellValue(sumarr[j-1]);
					}
				}

				// 新建一输出文件流
				FileOutputStream fOut = new FileOutputStream(file);
				// 把相应的Excel 工作簿存盘
				workbook.write(fOut);
				fOut.flush();
				// 操作结束，关闭文件
				fOut.close();
			}
			is.close();
			workbook.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
