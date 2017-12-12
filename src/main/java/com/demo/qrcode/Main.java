package com.demo.qrcode;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.FileNotFoundException;
import java.io.IOException;


public class Main {
	public static void main(String[] args) {
		String url_prefix = "http://baidu.com";//想要加网址的前缀
		String qrcode_output_path = "D:\\mycode\\fyt2\\wx\\admin\\qrcode\\src\\main\\java\\com\\yueke100\\qrcode/qrcode_imags/";
		try {
			ExcelUtilWithXSSF.getExcelAsFile("C:\\Users\\JimKan\\Desktop\\qrcode\\src\\test\\小学英语人教课时二维码.xlsx", qrcode_output_path,
					url_prefix);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		}
	}
}
