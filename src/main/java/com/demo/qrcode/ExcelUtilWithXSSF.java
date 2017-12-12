package com.demo.qrcode;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.google.zxing.BarcodeFormat;
import com.google.zxing.EncodeHintType;
import com.google.zxing.MultiFormatWriter;
import com.google.zxing.WriterException;
import com.google.zxing.common.BitMatrix;

public class ExcelUtilWithXSSF {

	private static final int DEFAULT_MEASURE = 470;
	private static final int MEASURE_130 = 130;
	private static final int MIN_MEASURE = 102;
	private static final String JPG_EXTENSION_NAME = ".jpg";
	private static final String PNG_EXTENSION_NAME = ".png";
	private static final String JPG_IMAGE_TYPE = "jpg";
	private static final String PNG_IMAGE_TYPE = "png";
	private static final String BG_IMAGE = "G:/SpringBoot_Workspace/qrcode/src/main/java/com/yueke100/qrcode/背景.tif";



	/**
	 * 得到Excel，并解析内容 对2007及以上版本 使用XSSF解析
	 * 
	 * @param file
	 * @throws FileNotFoundException
	 * @throws IOException
	 * @throws InvalidFormatException
	 */
	public static void getExcelAsFile(String file, String path, String url_prefix)
			throws FileNotFoundException, IOException, InvalidFormatException {
		InputStream ins = null;
		Workbook wb = null;
		ins = new FileInputStream(new File(file));
		wb = WorkbookFactory.create(ins);
		ins.close();
		System.err.println("工作薄数量" + wb.getNumberOfSheets());
		for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) {
			// 3.得到Excel工作表对象
			Sheet sheet = wb.getSheetAt(sheetIndex);
			String sheetName = sheet.getSheetName();
			System.err.println(sheetName);
			// 总行数
			int trLength = sheet.getLastRowNum();
			System.err.println("总行数：" + trLength);
			// 4.得到Excel工作表的行
			Row row = sheet.getRow(0);
			// 总列数
			// int tdLength = row.getLastCellNum();
			// 5.得到Excel工作表指定行的单元格
			Cell cell = row.getCell((short) 1);

			Row columnHeader = sheet.getRow(1);
			// Iterator<Cell> cellIterator = columnHeader.cellIterator();
			int qrcodeColumnIndex = 3;
			for (Cell header : columnHeader) {
				System.err.println(header.getRichStringCellValue() + "" + header.getColumnIndex());
				System.err.println(cell.getCellStyle().getFillBackgroundColor());
				if ("二维码编号".equals(header.getRichStringCellValue())) {
					qrcodeColumnIndex = header.getColumnIndex();
				}

			}
			System.err.println("二维码编号列索引：" + qrcodeColumnIndex);
			Cell qrcodeUrlHeader = columnHeader.createCell(qrcodeColumnIndex + 1);
			qrcodeUrlHeader.setCellValue("二维码URL");
			// 6.得到单元格样式
			CellStyle cellStyle = qrcodeUrlHeader.getCellStyle();
			cellStyle.setFillBackgroundColor((short) 64);
			for (int i = 2; i <= trLength; i++) {
				// 得到Excel工作表的行
				Row row1 = sheet.getRow(i);

				if (null == row1) {
					continue;
				}
				System.err.println(qrcodeColumnIndex);
				Cell cell1 = row1.getCell(qrcodeColumnIndex);
				String qrcode = cell1.getStringCellValue();
				String content = url_prefix + qrcode + "/0/";
				Cell qrcodeUrl = row1.createCell(qrcodeColumnIndex + 1);
				qrcodeUrl.setCellValue(content);

				encodeQRCode(content, qrcode, DEFAULT_MEASURE, DEFAULT_MEASURE, //
						path + File.separator + sheetName + File.separator + JPG_IMAGE_TYPE + File.separator, //
						JPG_EXTENSION_NAME, JPG_IMAGE_TYPE);
				encodeQRCode(content, qrcode, DEFAULT_MEASURE, DEFAULT_MEASURE,
						path + File.separator + sheetName + File.separator + PNG_IMAGE_TYPE, //
						PNG_EXTENSION_NAME, PNG_IMAGE_TYPE);
				// 获得每一列中的值
				System.out.print(cell1.getStringCellValue() + "                   ");
				System.out.println();
			}
		}
		

		// 将修改后的数据保存
		OutputStream out = new FileOutputStream(file);
		wb.write(out);
	}

	private static void encodeQRCode(String content, String imgName, int width, int height, String path,
			String extensionName, String imageType) {
		File file = new File(path);
		if (!file.exists()) {
			file.mkdirs();
		}
		MultiFormatWriter multiFormatWriter = new MultiFormatWriter();
		Map<EncodeHintType, String> hints = new HashMap<EncodeHintType, String>();
		hints.put(EncodeHintType.CHARACTER_SET, "UTF-8");
		BitMatrix bitMatrix;
		try {
			bitMatrix = multiFormatWriter.encode(content, BarcodeFormat.QR_CODE, width, height, hints);
			File file1 = new File(path, imgName + extensionName);
			MatrixToImageWriter.writeToFile(bitMatrix, imageType, file1);
		} catch (WriterException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} 
	}

	/**
	 * 创建Excel，并写入内容
	 */
	public static void CreateExcel() {

		// 1.创建Excel工作薄对象
		HSSFWorkbook wb = new HSSFWorkbook();
		// 2.创建Excel工作表对象
		HSSFSheet sheet = wb.createSheet("new Sheet");
		// 3.创建Excel工作表的行
		HSSFRow row = sheet.createRow(6);
		// 4.创建单元格样式
		CellStyle cellStyle = wb.createCellStyle();
		// 设置这些样式
		cellStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

		// 5.创建Excel工作表指定行的单元格
		row.createCell(0).setCellStyle(cellStyle);
		// 6.设置Excel工作表的值
		row.createCell(0).setCellValue("aaaa");

		row.createCell(1).setCellStyle(cellStyle);
		row.createCell(1).setCellValue("bbbb");

		// 设置sheet名称和单元格内容
		wb.setSheetName(0, "第一张工作表");
		// 设置单元格内容 cell.setCellValue("单元格内容");

		// 最后一步，将文件存到指定位置
		try {
			FileOutputStream fout = new FileOutputStream("E:/students.xls");
			wb.write(fout);
			fout.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// /**
	// * 创建Excel的实例
	// * @throws ParseException
	// */
	// public static void CreateExcelDemo1() throws ParseException{
	// List list = new ArrayList();
	// SimpleDateFormat df = new SimpleDateFormat("yyyy-mm-dd");
	// Student user1 = new Student(1, "张三", 16,true, df.parse("1997-03-12"));
	// Student user2 = new Student(2, "李四", 17,true, df.parse("1996-08-12"));
	// Student user3 = new Student(3, "王五", 26,false, df.parse("1985-11-12"));
	// list.add(user1);
	// list.add(user2);
	// list.add(user3);
	//
	//
	// // 第一步，创建一个webbook，对应一个Excel文件
	// HSSFWorkbook wb = new HSSFWorkbook();
	// // 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
	// HSSFSheet sheet = wb.createSheet("学生表一");
	// // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short
	// HSSFRow row = sheet.createRow((int) 0);
	// // 第四步，创建单元格，并设置值表头 设置表头居中
	// HSSFCellStyle style = wb.createCellStyle();
	// style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式
	//
	// HSSFCell cell = row.createCell((short) 0);
	// cell.setCellValue("学号");
	// cell.setCellStyle(style);
	// cell = row.createCell((short) 1);
	// cell.setCellValue("姓名");
	// cell.setCellStyle(style);
	// cell = row.createCell((short) 2);
	// cell.setCellValue("年龄");
	// cell.setCellStyle(style);
	// cell = row.createCell((short) 3);
	// cell.setCellValue("性别");
	// cell.setCellStyle(style);
	// cell = row.createCell((short) 4);
	// cell.setCellValue("生日");
	// cell.setCellStyle(style);
	//
	// // 第五步，写入实体数据 实际应用中这些数据从数据库得到，
	//
	// for (int i = 0; i < list.size(); i++)
	// {
	// row = sheet.createRow((int) i + 1);
	// Student stu = (Student) list.get(i);
	// // 第四步，创建单元格，并设置值
	// row.createCell((short) 0).setCellValue((double) stu.getId());
	// row.createCell((short) 1).setCellValue(stu.getName());
	// row.createCell((short) 2).setCellValue((double) stu.getAge());
	// row.createCell((short)3).setCellValue(stu.getSex()==true?"男":"女");
	// cell = row.createCell((short) 4);
	// cell.setCellValue(new SimpleDateFormat("yyyy-mm-dd").format(stu
	// .getBirthday()));
	// }
	// // 第六步，将文件存到指定位置
	// try
	// {
	// FileOutputStream fout = new FileOutputStream("E:/students.xls");
	// wb.write(fout);
	// fout.close();
	// }
	// catch (Exception e)
	// {
	// e.printStackTrace();
	// }
	//
	//
	//
	// }
}