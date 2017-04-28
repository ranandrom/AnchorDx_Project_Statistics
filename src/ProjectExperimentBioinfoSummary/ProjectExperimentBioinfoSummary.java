package ProjectExperimentBioinfoSummary;

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

import ProjectStatistics.ProjectStatistics;


/**
 * Description: <br/>
 * Function: This class is main class. <br/>
 * File Name: AnchorDx_Project_Statistics.java <br/>
 * Date: 2017-04-25
 * 
 * @author Luzhirong ramandrom@139.com
 * @version V1.0.0
 */
public class ProjectExperimentBioinfoSummary
{

	/**
	 * main方法，程序的入口.
	 * 
	 * @param args
	 * @throws Exception
	 */
	public static void main(String[] args)
	{
		// TODO Auto-generated method stub
		System.out.println();
		Calendar now_star = Calendar.getInstance();
		SimpleDateFormat formatter_star = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		System.out.println("程序开始时间: "+formatter_star.format(now_star.getTime()));
		System.out.println("===============================================");
		System.out.println("AnchorDx_Project_Statistics V1.0.0");
		System.out.println("***********************************************");
		System.out.println();

		SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd"); // 定义日期格式
		String ToDay = formatter.format(now_star.getTime()); // 获取当天日期

		String system = "windows"; // 操作系统
		String Output_Path = null;
		String Bioinfo_Path = null;
		String Experiment_Path = null;		
		
		int args_len = args.length; // 系统传入主函数的参数长度
		int logo = 0; // "-o"参数输入次数计算标志
		int logb = 0; // "-m"参数输入次数计算标志
		int loge = 0; // "-E"参数输入次数计算标志
		int logs = 0; // "-E"参数输入次数计算标志
		for (int len = 0; len < args_len; len += 2) {
			if (args[len].equals("-O") || args[len].equals("-o")) {	
				Output_Path = args[len + 1];
				logo++;
			} else if (args[len].equals("-B") || args[len].equals("-b")) {
				Bioinfo_Path = args[len + 1];
				logb++;
			} else if (args[len].equals("-E") || args[len].equals("-e")) {
				Experiment_Path = args[len + 1];
				loge++;
			} else if (args[len].equals("-S") || args[len].equals("-s")) {
				system = args[len + 1];
				logs++;
			} else if ((args_len == 1) && args[0].equals("-help")) {
				System.out.println();
				System.out.println("Version: V1.0.0");
				System.out.println();
				System.out.println("Usage:\t java -jar AnchorDx_Project_Statistics.jar [Options] [args...]");
				System.out.println();
				System.out.println("Options:");
				System.out.println("-help\t\t Obtain parameter description.");
				System.out.println("-O or -o\t Set output file. The default value is \"./项目实验生信汇总表\"");
				System.out.println("-B or -b\t Set Bioinfo summary table file path. The default value is \"/wdmycloud/anchordx_cloud/杨莹莹/项目进展汇总表\"");
				System.out.println("-E or -e\t Set progress of experimental items file path. The default value is \"/wdmycloud/anchordx_cloud/杨莹莹/项目进展汇总表\"");
				System.out.println("-S or -s\t Set running operating system . The default value is \"windows\"");
				System.out.println();
				return;
			} else {
				System.out.println();
				System.out.println("对不起，您输入的Options不存在，或者缺少所需参数，请参照以下参数提示输入！");
				System.out.println();
				System.out.println("Options:");
				System.out.println("-help\t\t Obtain parameter description.");
				System.out.println("-O or -o\t Set output file. The default value is \"./项目实验生信汇总表\"");
				System.out.println("-B or -b\t Set Bioinfo summary table file path. The default value is \"/wdmycloud/anchordx_cloud/杨莹莹/项目进展汇总表\"");
				System.out.println("-E or -e\t Set progress of experimental items file path. The default value is \"/wdmycloud/anchordx_cloud/杨莹莹/项目进展汇总表\"");
				System.out.println("-S or -s\t Set running operating system . The default value is \"windows\"");
				System.out.println();
				return;
			}
			if (logo > 1 || logb > 1 || loge > 1 || logs > 1) {
				System.out.println();
				System.out.println("对不起，您输的入Options有重复，请参照以下参数提示输入！");
				System.out.println();
				System.out.println("Options:");
				System.out.println("-help\t\t Obtain parameter description.");
				System.out.println("-O or -o\t Set output file. The default value is \"./项目实验生信汇总表\"");
				System.out.println("-B or -b\t Set Bioinfo summary table file path. The default value is \"/wdmycloud/anchordx_cloud/杨莹莹/项目进展汇总表\"");
				System.out.println("-E or -e\t Set progress of experimental items file path. The default value is \"/wdmycloud/anchordx_cloud/杨莹莹/项目进展汇总表\"");
				System.out.println("-S or -s\t Set running operating system . The default value is \"windows\"");
				System.out.println();
				return;
			}
		}
		
		if (system.equals("linux")) {
			if (logo == 0) {
				Output_Path = "./项目实验生信汇总表";
			}
			if (logb == 0) {
				Bioinfo_Path = "/wdmycloud/anchordx_cloud/杨莹莹/项目-生信-汇总表";
			} 
			if (loge == 0) {
				Experiment_Path = "/wdmycloud/anchordx_cloud/杨莹莹/项目进展汇总表";
			}
		} else {
			if (logo == 0) {
				Output_Path = "./项目实验生信汇总表";
			}
			if (logb == 0) {
				Bioinfo_Path = "\\\\wdmycloud\\anchordx_cloud\\杨莹莹\\项目-生信-汇总表";
			}
			if (loge == 0) {
				Experiment_Path = "\\\\wdmycloud\\anchordx_cloud\\杨莹莹\\项目进展汇总表";
			}
		}
		String Output_dir_name = null;
		if (system.equals("linux")) {
			Output_dir_name = Output_Path + "/" + ToDay;
		} else {
			Output_dir_name = Output_Path + "\\" + ToDay;
		}
		ProjectStatistics.myMkdir(Output_dir_name); // 创建输出目录
		
		String Bioinfo_File = null;
		if (logb == 0) {
			String Newest_Dir = ProjectStatistics.getNewestFileDir(Bioinfo_Path); // 获取最新文件所在目录名			
			if (system.equals("linux")) {
				Bioinfo_File = Bioinfo_Path + "/" + Newest_Dir; // 生信带路径文件名
			} else {
				Bioinfo_File = Bioinfo_Path + "\\" + Newest_Dir; // 生信带路径文件名
			}
		} else {
			Bioinfo_File = Bioinfo_Path; // 生信带路径文件名
		}
		
		File Bio_file = new File(Bioinfo_File);
		ArrayList<String> Bioinfo_list = new ArrayList<String>();
		searchBioinfoFile(Bio_file, Bioinfo_list, system); // 获取最新生信文件列表
		
		File Exp_file = new File(Experiment_Path);
		ArrayList<String> Experiment_list = new ArrayList<String>();
		searchExperimentFile(Exp_file, Experiment_list, system); // 获取需要合并的实验文件列表
		
		String Bioinfo_Plasma_FilePath = null;
		String Bioinfo_Tissue_FilePath = null;
		String Bioinfo_BC_FilePath = null;
		String OutputFilePath = null;
		
		ArrayList<String> ExperimentHeadlist = new ArrayList<String>(); // 头信息列表
		ArrayList<ArrayList<String>> Experiment_Data_list = new ArrayList<ArrayList<String>>(); // 实验文件数据列表
		ArrayList<String> Bioinfo_Plasma_Data_list = new ArrayList<String>(); // 生信血浆文件数据列表
		ArrayList<String> Bioinfo_Tissue_Data_list = new ArrayList<String>(); // 生信组织文件数据列表
		ArrayList<String> Bioinfo_BC_Data_list = new ArrayList<String>(); // 生信白细胞文件数据列表
		
		//循环合并每个实验表
		for(int i = 0; i  < Experiment_list.size(); i++){
			String Experiment_str[] = Experiment_list.get(i).split("\t");			
			File Experiment_file = new File(Experiment_str[0]);
			ExperimentHeadlist.clear();
			ExperimentHeadlist = readExperimentHead(Experiment_str[0]); // 返回带sheet名的头信息，以制表符分隔
			if (system.equals("linux")) {
				OutputFilePath = Output_dir_name + "/" + Experiment_str[1] + "_项目实验生信汇总表_" + ToDay + ".xlsx";
			} else {
				OutputFilePath = Output_dir_name + "\\" + Experiment_str[1] + "_项目实验生信汇总表_" + ToDay + ".xlsx";
			}
			createXlsx(new File(OutputFilePath), ExperimentHeadlist); // 创建合并文件，表的个数与 Experiment_file 里的表的个数相同。			
			Bioinfo_Plasma_Data_list.clear();
			Bioinfo_Tissue_Data_list.clear();
			Bioinfo_BC_Data_list.clear();
			Experiment_Data_list.clear();
			Experiment_Data_list = readExperimentXlsx(Experiment_file); // 读实验表数据
			
			for (int j = 0;j  < Bioinfo_list.size(); j++) {
				String Bioinfo_str[] = Bioinfo_list.get(j).split("\t");
				if (Experiment_str[1].equals(Bioinfo_str[2])) {					
					if (Bioinfo_str[1].equals("Plasma")) {
						Bioinfo_Plasma_FilePath = Bioinfo_str[0];
						File Bioinfo_Plasma_file = new File(Bioinfo_Plasma_FilePath);
						Bioinfo_Plasma_Data_list = readBioinfoXlsx(Bioinfo_Plasma_file); // 读血浆表数据
					} else if (Bioinfo_str[1].equals("Tissue")) {
						Bioinfo_Tissue_FilePath = Bioinfo_str[0];
						File Bioinfo_Tissue_file = new File(Bioinfo_Tissue_FilePath);
						Bioinfo_Tissue_Data_list = readBioinfoXlsx(Bioinfo_Tissue_file); // 读组织表数据
					} else if (Bioinfo_str[1].equals("BC")) {
						Bioinfo_BC_FilePath = Bioinfo_str[0];
						File Bioinfo_BC_file = new File(Bioinfo_BC_FilePath);
						Bioinfo_BC_Data_list = readBioinfoXlsx(Bioinfo_BC_file); // 读白细胞表数据
					}
				}
			}
			mergeFile(OutputFilePath, Experiment_Data_list, Bioinfo_Plasma_Data_list, Bioinfo_Tissue_Data_list, Bioinfo_BC_Data_list); // 合并文件
			System.out.println(Experiment_str[1] + "项目已合并完成！");
		}
		
		ProjectStatistics.projectStatisticsMain(Output_dir_name, system);
		System.out.println();
		System.out.println("项目统计已完成！");
		
		Calendar now_end = Calendar.getInstance();
		SimpleDateFormat formatter_end = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		System.out.println();
		System.out.println("==============================================");
		System.out.println("程序结束时间: "+formatter_end.format(now_end.getTime()));
		System.out.println();

	}
	
	/**
	 * 查找生信文件
	 * @param des_file
	 * @param list
	 */
	public static void searchBioinfoFile(File des_file, ArrayList<String> list, String system)
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
					
					if( !(FileName.startsWith("~$")) && !(FileName.startsWith("."))  ){
						String FileName_str[] = FileName.split("_");
						String str = null;
						if (system.equals("linux")) {
							str = Folder + "/" + FileName + "\t" + FileName_str[0]+ "\t" + FileName_str[1];
						} else {
							str = Folder + "\\" + FileName + "\t" + FileName_str[0]+ "\t" + FileName_str[1];
						}
						list.add(str);
					}
					continue;
				}else{
					//如果是目录，则递归
					searchBioinfoFile(pathname, list, system);
				}
			}
		}catch(Exception e){
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }
	
	/**
	 * 查找实验文件
	 * @param des_file
	 * @param list
	 */
	public static void searchExperimentFile(File des_file, ArrayList<String> list, String system)
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
					
					if( !(FileName.startsWith("~$")) && !(FileName.startsWith("."))  ){
						String FileName_str[] = FileName.split("_");
						if(FileName_str[1].contains("项目进展")){
							String str = null;
							if (system.equals("linux")) {
								str = Folder + "/" + FileName + "\t" + FileName_str[0];
							} else {
								str = Folder + "\\" + FileName + "\t" + FileName_str[0];
							}
							list.add(str);
						}
					}
					continue;
				}else{
					//如果是目录，则递归
					searchExperimentFile(pathname, list, system);
				}
			}
		}catch(Exception e){
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }
	
	/**
	 * 读实验文件，返回标头列表（以 '\t' 合并，前面带该表格的表格名的 String）的列表。
	 * @param Experiment_str
	 * @return
	 */
	public static ArrayList<String> readExperimentHead(String Experiment_str)
	{		
		ArrayList<String> Data_list = new ArrayList<String>();
		File file = new File(Experiment_str);
		try {
			//System.out.println(Experiment_str + "readHead1");
			InputStream is = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(is);
			XSSFSheet sheet = null;
			int Sheet_Num = wb.getNumberOfSheets(); // 获取工作薄个数
			//System.out.println(Experiment_str + "readHead");
			
			String data = null;
			for(int numSheet = 0; numSheet < Sheet_Num; numSheet++ ){
				sheet = wb.getSheetAt(numSheet); // 获取工作薄
				String Sheet_Name = sheet.getSheetName(); // 获取当前工作薄名字
				//System.out.println(Sheet_Name);
				XSSFRow xssfrow = sheet.getRow(0);

				try {
					XSSFCell xssfcell0 = xssfrow.getCell(0);
				}  catch (Exception e) {
					// TODO Auto-generated catch block
					//e.printStackTrace();
					//System.out.println(Experiment_str + "'s " + Sheet_Name + ", this sheet is null!");
					continue;
				}
				
				int headlog = 0;
				// 获取当前工作薄 "TNM stage" 列前的每一列。
				for (int j = xssfrow.getFirstCellNum(); j < xssfrow.getLastCellNum(); j++) {
					//System.out.println(xssfrow.getLastCellNum());
					XSSFCell xssfcell = xssfrow.getCell(j);
					if(j == 0){
						//data = xssfcell.getStringCellValue().trim();
						int xcellty = xssfcell.getCellType();
						if( xcellty == 0 ){
							if(HSSFDateUtil.isCellDateFormatted(xssfcell)){//判断是否为日期格式
								Date date = xssfcell.getDateCellValue();
								SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
								data = dateFormat.format(date);
							}else{
								xssfcell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);//设置单元格为数值类型
								HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
								String cellFormatted = dataFormatter.formatCellValue(xssfcell);//按单元格数据格式类型读取数据
								data = cellFormatted;
							}
						}else{
							xssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
							data = xssfcell.getStringCellValue().trim();
						}
					}else{
						//data += "\t" + xssfcell.getStringCellValue().trim();
						int xcellty = xssfcell.getCellType();
						if( xcellty == 0 ){
							if(HSSFDateUtil.isCellDateFormatted(xssfcell)){//判断是否为日期格式
								Date date = xssfcell.getDateCellValue();
								SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
								data += "\t" + dateFormat.format(date);
							}else{
								xssfcell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);//设置单元格为数值类型
								HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
								String cellFormatted = dataFormatter.formatCellValue(xssfcell);//按单元格数据格式类型读取数据
								data += "\t" + cellFormatted;
							}
						}else{
							xssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
							data += "\t" + xssfcell.getStringCellValue().trim();
						}
					}
					int xcellty = xssfcell.getCellType();
					String head = null;
					if( xcellty == 0 ){
						if(HSSFDateUtil.isCellDateFormatted(xssfcell)){//判断是否为日期格式
							Date date = xssfcell.getDateCellValue();
							SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
							head = dateFormat.format(date);
						}else{
							xssfcell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);//设置单元格为数值类型
							HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
							String cellFormatted = dataFormatter.formatCellValue(xssfcell);//按单元格数据格式类型读取数据
							head = cellFormatted;
						}
					}else{
						xssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
						head = xssfcell.getStringCellValue().trim();
					}
					if(head.equals("TNM stage") ){
						headlog = 1;
						break;
					}
				}
				data = Sheet_Name +  "\t" + data;
				if (headlog == 1) {
					Data_list.add(data);
				}
			}
			is.close();
			wb.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return Data_list;
	}
	
	//新建合并文件
	public static void createXlsx(File file, ArrayList<String> ExperimentHeadlist)
	{
		try{
			XSSFWorkbook workbook = new XSSFWorkbook();
			for(int j = 0; j < ExperimentHeadlist.size(); j++){
				String str_head[] = ExperimentHeadlist.get(j).split("\t");
				workbook.createSheet(str_head[0]);
			}
			for(int j = 0; j < ExperimentHeadlist.size(); j++){
				String str_h[] = ExperimentHeadlist.get(j).split("\t");
				// 创建Excel的工作sheet,对应到一个excel文档的tab  
				XSSFSheet sheet = workbook.getSheet(str_h[0]);	//获取工作薄;
				// 在索引0的位置创建行（最顶端的行）
				XSSFRow row0 = sheet.createRow((short) 0);
				
				String head_row0 = "Sample ID"+"\t"+
				"Pre-lib name"+"\t"+
						"Identification name"+"\t"+
				"Sequencing info"+"\t"+
						"Sequencing file name"+"\t"+
				"Mapping%"+"\t"+
						"Total PF reads"+"\t"+
				"Mean_insert_size"+"\t"+
						"Median_insert_size"+"\t"+
				"On target%"+"\t"+
						"Pre-dedup mean bait coverage"+"\t"+
				"Deduped mean bait coverage"+"\t"+
						"Deduped mean target coverage"+"\t"+
				"% target bases > 30X"+"\t"+
						"Uniformity (0.2X mean)"+"\t"+
				"C methylated in CHG context"+"\t"+
						"C methylated in CHH context"+"\t"+
				"C methylated in CpG context"+"\t"+
						"QC result"+"\t"+
				"Date of QC"+"\t"+
						"Path to sorted.deduped.bam"+"\t"+
				"Date of path update"+"\t"+
						"Bait set"+"\t"+
				"Sample QC"+"\t"+
						"Failed QC Detail"+"\t"+
				"Warning QC Detail"+"\t"+
						"Check"+"\t"+
				"Note1"+"\t"+
						"Note2"+"\t"+
				"Note3";
				
				//1、创建字体，设置其为红色：
				XSSFFont font = workbook.createFont();
				font.setColor(HSSFFont.COLOR_RED);
				font.setFontHeightInPoints((short)10);
				font.setFontName("Palatino Linotype");
				//font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
				//2、创建格式
				XSSFCellStyle cellStyle= workbook.createCellStyle();
				cellStyle.setFont(font);
				cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
				cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
				cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
				cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
				cellStyle.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
				cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				 
				//1、创建字体，设置其为粗体，背景蓝色：
				XSSFFont font1 = workbook.createFont();
				//font1.setColor(HSSFFont.COLOR_RED);
				font1.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
				font1.setFontHeightInPoints((short)10);
				font1.setFontName("Palatino Linotype");
				//2、创建格式
				XSSFCellStyle cellStyle1= workbook.createCellStyle();
				cellStyle1.setFont(font1);
				cellStyle1.setBorderTop(XSSFCellStyle.BORDER_THIN);
				cellStyle1.setBorderBottom(XSSFCellStyle.BORDER_THIN);
				cellStyle1.setBorderLeft(XSSFCellStyle.BORDER_THIN);
				cellStyle1.setBorderRight(XSSFCellStyle.BORDER_THIN);
				cellStyle1.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
				cellStyle1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				 
				//1、创建字体，设置其为红色、粗体，背景绿色：
				XSSFFont font2 = workbook.createFont();
				font2.setColor(HSSFFont.COLOR_RED);
				font2.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
				font2.setFontHeightInPoints((short)10);
				font2.setFontName("Palatino Linotype");
				//2、创建格式
				XSSFCellStyle cellStyle2= workbook.createCellStyle();
				cellStyle2.setFont(font2);
				cellStyle2.setBorderTop(XSSFCellStyle.BORDER_THIN);
				cellStyle2.setBorderBottom(XSSFCellStyle.BORDER_THIN);
				cellStyle2.setBorderLeft(XSSFCellStyle.BORDER_THIN);
				cellStyle2.setBorderRight(XSSFCellStyle.BORDER_THIN);
				cellStyle2.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
				cellStyle2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				 
				//1、创建字体大小为10，背景蓝色：
				XSSFFont font3 = workbook.createFont();
				font3.setFontHeightInPoints((short)10);
				font3.setFontName("Palatino Linotype");
				//2、创建格式
				XSSFCellStyle cellStyle3= workbook.createCellStyle();
				cellStyle3.setFont(font3);
				cellStyle3.setBorderTop(XSSFCellStyle.BORDER_THIN);
				cellStyle3.setBorderBottom(XSSFCellStyle.BORDER_THIN);
				cellStyle3.setBorderLeft(XSSFCellStyle.BORDER_THIN);
				cellStyle3.setBorderRight(XSSFCellStyle.BORDER_THIN);
				cellStyle3.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
				cellStyle3.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				 
				//1、创建字体大小为10，背景黄色：
				XSSFFont font4 = workbook.createFont();
				font4.setFontHeightInPoints((short)10);
				font4.setFontName("Palatino Linotype");
				//2、创建格式
				XSSFCellStyle cellStyle4= workbook.createCellStyle();
				cellStyle4.setFont(font4);
				cellStyle4.setBorderTop(XSSFCellStyle.BORDER_THIN);
				cellStyle4.setBorderBottom(XSSFCellStyle.BORDER_THIN);
				cellStyle4.setBorderLeft(XSSFCellStyle.BORDER_THIN);
				cellStyle4.setBorderRight(XSSFCellStyle.BORDER_THIN);
				cellStyle4.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
				cellStyle4.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				 
				//1、创建字体，设置其为粗体，背景黄色：
				XSSFFont font5 = workbook.createFont();
				//font1.setColor(HSSFFont.COLOR_RED);
				font5.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
				font5.setFontHeightInPoints((short)10);
				font5.setFontName("Palatino Linotype");
				//2、创建格式
				XSSFCellStyle cellStyle5= workbook.createCellStyle();
				cellStyle5.setFont(font5);
				cellStyle5.setBorderTop(XSSFCellStyle.BORDER_THIN);
				cellStyle5.setBorderBottom(XSSFCellStyle.BORDER_THIN);
				cellStyle5.setBorderLeft(XSSFCellStyle.BORDER_THIN);
				cellStyle5.setBorderRight(XSSFCellStyle.BORDER_THIN);
				cellStyle5.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
				cellStyle5.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				
				//1、创建字体，设置其为粗体，背景橘色：
				XSSFFont font6 = workbook.createFont();
				//font6.setColor(HSSFFont.COLOR_RED);
				font6.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
				font6.setFontHeightInPoints((short)10);
				font6.setFontName("Palatino Linotype");
				//2、创建格式
				XSSFCellStyle cellStyle6 = workbook.createCellStyle();
				cellStyle6.setFont(font6);
				cellStyle6.setBorderTop(XSSFCellStyle.BORDER_THIN);
				cellStyle6.setBorderBottom(XSSFCellStyle.BORDER_THIN);
				cellStyle6.setBorderLeft(XSSFCellStyle.BORDER_THIN);
				cellStyle6.setBorderRight(XSSFCellStyle.BORDER_THIN);
				cellStyle6.setFillForegroundColor(HSSFColor.TAN.index);
				cellStyle6.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				
				//1、创建字体，设置其为粗体，红字，背景橘色：
				XSSFFont font7 = workbook.createFont();
				font7.setColor(HSSFFont.COLOR_RED);
				font7.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
				font7.setFontHeightInPoints((short)10);
				font7.setFontName("Palatino Linotype");
				//2、创建格式
				XSSFCellStyle cellStyle7 = workbook.createCellStyle();
				cellStyle7.setFont(font7);
				cellStyle7.setBorderTop(XSSFCellStyle.BORDER_THIN);
				cellStyle7.setBorderBottom(XSSFCellStyle.BORDER_THIN);
				cellStyle7.setBorderLeft(XSSFCellStyle.BORDER_THIN);
				cellStyle7.setBorderRight(XSSFCellStyle.BORDER_THIN);
				cellStyle7.setFillForegroundColor(HSSFColor.TAN.index);
				cellStyle7.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
								 
				String str_head = ExperimentHeadlist.get(j) + "\t" + head_row0;
				String str_Data[] = ExperimentHeadlist.get(j).split("\t");
				String str_head_row0[] = str_head.split("\t");
				// 在单元格中输入一些内容
				for(int i = 1; i < str_head_row0.length; i++ ){
					// 在索引0的位置创建单元格（左上端）
					XSSFCell cell = row0.createCell(i-1);
					if( i < 4 ){// 实验表格的 "Sample ID" ～ "Sequencing info"：红字橘底
						cell.setCellValue(str_head_row0[i]);
						cell.setCellStyle(cellStyle7);
					}else if( i >= str_Data.length && i < str_Data.length+3 ){// 生信表格的 "Sample ID" ～ "Sequencing info"：红字绿底。
						cell.setCellValue(str_head_row0[i]);
						cell.setCellStyle(cellStyle2);
					}else if( i == str_head_row0.length-10 || i == str_head_row0.length-9 ){// "Path to sorted.deduped.bam"、"Date of path update"：黑字黄底。
						cell.setCellStyle(cellStyle5);
						cell.setCellValue(str_head_row0[i]);
					}else if(i > str_Data.length){// 剩下的生信表格的列：黑字蓝底。
						cell.setCellStyle(cellStyle1);
						cell.setCellValue(str_head_row0[i]);
					}else{// 剩下的部分（实验表格的列）：黑子橘底。
						cell.setCellStyle(cellStyle6);
						cell.setCellValue(str_head_row0[i]);
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
			workbook.close();
		}catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	/**
	 * 读实验文件数据
	 * @param file
	 * @return
	 */
	public static ArrayList<ArrayList<String>> readExperimentXlsx(File file)
	{	
		ArrayList< ArrayList<String> > Data_list = new ArrayList< ArrayList<String> >();
		try {
		InputStream is = new FileInputStream(file);
		XSSFWorkbook wb = new XSSFWorkbook(is);
		//XSSFSheet sheet = wb.getSheetAt(2);	//获取第三个工作薄
		XSSFSheet sheet = null;
		int Sheet_Num = wb.getNumberOfSheets();//获取工作薄个数
		//System.out.println(Sheet_Num);
		
		String data = null;
		int celllog = 0;// 读取的最后一列。
		for(int numSheet = 0; numSheet < Sheet_Num; numSheet++ ){
			ArrayList<String> datalist = new ArrayList<String>();
			sheet = wb.getSheetAt(numSheet);	//获取工作薄
			String Sheet_Name = sheet.getSheetName();//获取当前工作薄名字
			//System.out.println(Sheet_Name);
			
			XSSFRow xssfrow0 = sheet.getRow(0);
			
			try {
				XSSFCell xssfcell0 = xssfrow0.getCell(0);
			}  catch (Exception e) {
				// TODO Auto-generated catch block
				//e.printStackTrace();
				//System.out.println(Sheet_Name + "\tnull");
				continue;
			}
			// 获取当前工作薄 "TNM stage" 列前的每一列。
			for (int j = xssfrow0.getFirstCellNum(); j < xssfrow0.getLastCellNum(); j++) {
				XSSFCell xssfcell = xssfrow0.getCell(j);
				int xcellty = xssfcell.getCellType();
				String head = null;
				if( xcellty == 0 ){
					if(HSSFDateUtil.isCellDateFormatted(xssfcell)){//判断是否为日期格式
						Date date = xssfcell.getDateCellValue();
						SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
						head = dateFormat.format(date);
					}else{
						xssfcell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);//设置单元格为数值类型
						HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
						String cellFormatted = dataFormatter.formatCellValue(xssfcell);//按单元格数据格式类型读取数据
						head = cellFormatted;
					}
				}else{
					xssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
					head = xssfcell.getStringCellValue().trim();
				}
				if( head.equals("TNM stage") ){
					celllog = j; 
					break;
				}
			}
			int nullrowlog = 0;
			// 获取当前工作薄的每一行
			for (int i = sheet.getFirstRowNum()+1; i <= sheet.getLastRowNum(); i++) {
				int nulllog = 0;
				XSSFRow xssfrow = sheet.getRow(i);
				//System.out.println(i);
				
				if ( xssfrow == null || (checkRowNull(xssfrow) == 0) ) {
					for(int j = 0; j < celllog; j++){
						if(j == 0){
							data = "null";
						}else{
							data += "\t" + "null";
						}
					}
					nulllog++;
					nullrowlog++;
				}else{
					// 获取当前工作薄的每一列
					for (int j = 0; j <= celllog; j++) {
						XSSFCell xssfcell = xssfrow.getCell(j);
						
						if(j == 0 ){
							if( xssfcell == null  || (("") == String.valueOf(xssfcell)) || xssfcell.toString().equals("") || xssfcell.getCellType() == HSSFCell.CELL_TYPE_BLANK ){
								data = "null";
							}else{
								int xcellty = xssfcell.getCellType();
								if( xcellty == 0 ){
									if(HSSFDateUtil.isCellDateFormatted(xssfcell)){//判断是否为日期格式
										Date date = xssfcell.getDateCellValue();
										SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
										data = dateFormat.format(date);	//以日期格式获取数据
									}else{
										xssfcell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);//设置单元格为数值类型
										HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
										String cellFormatted = dataFormatter.formatCellValue(xssfcell);//按单元格数据格式类型读取数据
										data = cellFormatted;
									}
								}else{
									xssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
									data = xssfcell.getStringCellValue().trim();
								}
								nulllog++;
							}
							//System.out.println(data);
						}else{
							if( xssfcell == null  || (("") == String.valueOf(xssfcell)) || xssfcell.toString().equals("") || xssfcell.getCellType() == HSSFCell.CELL_TYPE_BLANK ){
								data += "\t" +  "null";
							}else{
								int xcellty = xssfcell.getCellType();
								if( xcellty == 0 ){
									if(HSSFDateUtil.isCellDateFormatted(xssfcell)){//判断是否为日期格式
										Date date = xssfcell.getDateCellValue();
										SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
										data += "\t" + dateFormat.format(date);
									}else{
										xssfcell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);//设置单元格为数值类型
										HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
										String cellFormatted = dataFormatter.formatCellValue(xssfcell);//按单元格数据格式类型读取数据
										data += "\t" + cellFormatted;
									}
								}else{
									xssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
									data += "\t" + xssfcell.getStringCellValue().trim();
								}
								nulllog++;
							}							
						}
					}
				}
				if(nulllog != 0){
					datalist.add(data);
					//System.out.println(data);
				}
				if(nullrowlog > 5){
					for(int n = 0; n <= 5; n++){
						datalist.remove(datalist.size()-1);
					}
					break;
				}
			}
			Data_list.add(datalist);
		}
			is.close();
			wb.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return Data_list;
	}

	/**
	 * 判断行为空,如果为空，则返回0
	 * @param xssfRow
	 * @return
	 */
	public static int checkRowNull(XSSFRow xssfRow)
	{
		int num = 0;
		// 获取当前工作薄的每一列
		for (int j = xssfRow.getFirstCellNum(); j < xssfRow.getLastCellNum(); j++) {
			XSSFCell xssfcell = xssfRow.getCell(j);
			//String cellValue = String.valueOf(xssfcell);
			if( xssfcell == null  || (("") == String.valueOf(xssfcell)) || xssfcell.toString().equals("") || xssfcell.getCellType() == HSSFCell.CELL_TYPE_BLANK ){
				continue;
			}else{
				num++;
			}
		}
		return num;
	}
	
	/**
	 * 读生信文件数据
	 * @param file
	 * @return
	 */
	public static ArrayList<String> readBioinfoXlsx(File file)
	{
		ArrayList<String> Data_list = new ArrayList<String>();
		try{
			InputStream is = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(is);
			XSSFSheet sheet = wb.getSheetAt(0);	// 获取第1个工作薄
			String data = null;
			// 获取当前工作薄的每一行
			for (int i = sheet.getFirstRowNum()+1; i <= sheet.getLastRowNum(); i++) {
		
				XSSFRow xssfrow = sheet.getRow(i);
					
				// 获取当前工作薄的每一列
				for (int j = xssfrow.getFirstCellNum(); j < xssfrow.getLastCellNum(); j++) {
					XSSFCell xssfcell = xssfrow.getCell(j);
						
					if(j == 0){
						xssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
						data = xssfcell.getStringCellValue().trim();
					}else{
						xssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
						data += "\t" + xssfcell.getStringCellValue().trim();
					}
				}
				Data_list.add(data);
			}
			is.close();
			wb.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return Data_list;
	}
	
	/**
	 * 合并文件
	 * @param Outputfile
	 * @param Experiment_Data_list
	 * @param Bioinfo_Plasma_Data_list
	 * @param Bioinfo_Tissue_Data_list
	 * @param Bioinfo_BC_Data_list
	 */
	public static void mergeFile(String OutputFilePath, ArrayList<ArrayList<String>> Experiment_Data_list, ArrayList<String> Bioinfo_Plasma_Data_list,
									ArrayList<String> Bioinfo_Tissue_Data_list, ArrayList<String> Bioinfo_BC_Data_list)
	{
		ArrayList<String> Data_list = new ArrayList<String>();
		ArrayList<String> Bioinfo_Data = new ArrayList<String>();
		ArrayList<String> Bioinfo_remove = new ArrayList<String>();
		ArrayList<String> Bioinfo_Data_list = new ArrayList<String>();
		
		try{
			File OutPutfile = new File(OutputFilePath);
			FileInputStream is = new FileInputStream(OutPutfile);
			XSSFWorkbook workbook = new XSSFWorkbook(is);
			XSSFSheet sheet = null;
			int Sheet_Num = workbook.getNumberOfSheets();//获取工作薄个数
			
			for(int numSheet = 0; numSheet < Sheet_Num; numSheet++ ){
				sheet = workbook.getSheetAt(numSheet);	//获取工作薄
				String Sheet_Name = sheet.getSheetName();//获取当前工作薄名字
				for(int j = 0; j < Experiment_Data_list.get(numSheet).size(); j++){
					String Experiment[] = Experiment_Data_list.get(numSheet).get(j).split("\t");
					if( !(Experiment[1].equals("null")) ){
						String Pre_lib_name[] = Experiment[1].split("-");
						if(Pre_lib_name.length > 1){
							if( Pre_lib_name[Pre_lib_name.length-1].startsWith("PS") ){
								Bioinfo_Data_list = Bioinfo_Plasma_Data_list;
								break;
							}else if( Pre_lib_name[Pre_lib_name.length-1].startsWith("BC") ){
								Bioinfo_Data_list = Bioinfo_BC_Data_list;
								break;
							}else if( Pre_lib_name[Pre_lib_name.length-1].startsWith("F") ){
								Bioinfo_Data_list = Bioinfo_Tissue_Data_list;
								break;
							}else{
								continue;
							}
						}
					}
				}
				
				String null_data = null;
				if (Experiment_Data_list.get(numSheet).size() == 0) {
					continue;
				}
				String Experiment[] = Experiment_Data_list.get(numSheet).get(0).split("\t");
				for(int j = 0; j < Experiment.length; j++){
					if(j == 0){
						null_data = "null";
					}else{
						null_data += "\t" + "null";
					}
				}
				int rownum = 1;
				String chrB = null;
				String chrE = null;
				Bioinfo_Data.clear();
				Bioinfo_remove.clear();
				//写数据
				for(int j = 0; j < Experiment_Data_list.get(numSheet).size(); j++){

					String data = null;
					String strE = null;
					String strBprototype = null;
					String strBaddDNA = null;
					String strBrmDNA = null;
					int log = 0;
					String str_Experiment[] = Experiment_Data_list.get(numSheet).get(j).split("\t");
					Data_list.clear();
					for(int x = 0; x < Bioinfo_Data_list.size(); x++){
						String str_Bioinfo[] = Bioinfo_Data_list.get(x).split("\t");
						String str_Bio[] = str_Bioinfo[1].split("-");
						for(int y = 0; y < str_Bio.length-1; y++){
							if(y == 0){
								strBprototype = str_Bio[y];
								strBaddDNA = str_Bio[y];
							} else if (y == 1) {
								if (str_Bio[y].equals("DNA")) {
									strBprototype += "-" + str_Bio[y];
									strBaddDNA += "-" + str_Bio[y];
								} else {
									strBprototype += "-" + str_Bio[y];
									strBaddDNA += "-DNA-" + str_Bio[y];
									strBrmDNA += "-" + str_Bio[y];
								}
							} else {
								strBprototype += "-" + str_Bio[y];
								strBaddDNA += "-" + str_Bio[y];
								strBrmDNA += "-" + str_Bio[y];
							}
						}
						if(str_Experiment[1].contains(strBprototype) || str_Experiment[1].contains(strBaddDNA) || str_Experiment[1].contains(strBrmDNA)){
							if(log == 0){
								data = Experiment_Data_list.get(numSheet).get(j) + "\t" + Bioinfo_Data_list.get(x);
							} else {
								data = null_data + "\t" + Bioinfo_Data_list.get(x);
							}
							Data_list.add(data);
							Bioinfo_Data.add(Bioinfo_Data_list.get(x));
							log++;
							
							if (strE == null) {
								String str_Exp[] = str_Experiment[1].split("-");
								strE = str_Exp[2];
								for(int i = 0 ; i < strE.length(); i++){ //循环遍历字符串
									if(Character.isLetter(strE.charAt(i))){   //用char包装类中的判断字母的方法判断每一个字符
										chrE = String.valueOf(strE.charAt(i));
				                    }
								}
							}
						}else{
							continue;
						}
					}
					if (Data_list.size() != 0) {
						for(int x = 0; x < Data_list.size(); x++){
							XSSFRow row = sheet.createRow((short) rownum++);
							String str_row[] = Data_list.get(x).split("\t");
							for(int i = 0; i < str_row.length; i++ ){
								// 在索引0的位置创建单元格（左上端）
								XSSFCell cell = row.createCell(i);
								if(str_row[i].equals("null")){
									cell.setCellValue("");
								}else{
									cell.setCellValue(str_row[i]);
								}
							}
						}
					}else{
						XSSFRow row = sheet.createRow((short) rownum++);
						for(int i = 0; i < str_Experiment.length; i++ ){
							// 在索引0的位置创建单元格（左上端）
							XSSFCell cell = row.createCell(i);
							if(str_Experiment[i].equals("null")){
								cell.setCellValue("");
							}else{
								cell.setCellValue(str_Experiment[i]);
							}
						}
					}
				}
				for(int x = 0; x < Bioinfo_Data_list.size(); x++){
					if(!(Bioinfo_Data.contains(Bioinfo_Data_list.get(x)))){					
						if (chrE != null) { // 判断是否带有字母
							String str_Bioinfo[] = Bioinfo_Data_list.get(x).split("\t");
							String str_Bio[] = str_Bioinfo[1].split("-");
							for(int i = 0 ; i < str_Bio[2].length(); i++){ //循环遍历字符串
								if(Character.isLetter(str_Bio[2].charAt(i))){   //用char包装类中的判断字母的方法判断每一个字符
									chrB = String.valueOf(str_Bio[2].charAt(i));
			                    }
							}
							
							if (chrE.equals(chrB)) { // 判断是否有相同的字母
								Bioinfo_remove.add(Bioinfo_Data_list.get(x));
							} else {
								continue;
							}
						} else {
							Bioinfo_remove.add(Bioinfo_Data_list.get(x));
						}
					}
				}
				for(int x = 0; x < Bioinfo_remove.size(); x++){
					XSSFRow row = sheet.createRow((short) rownum++);
					String removedata = null_data + "\t" + Bioinfo_remove.get(x);
					String str_row[] = removedata.split("\t");
					for(int i = 0; i < str_row.length; i++ ){
						// 在索引0的位置创建单元格（左上端）
						XSSFCell cell = row.createCell(i);
						if(str_row[i].equals("null")){
							cell.setCellValue("");
						}else{
							cell.setCellValue(str_row[i]);
						}
					}
				}

				// 新建一输出文件流
				FileOutputStream fOut = new FileOutputStream(OutPutfile);
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
