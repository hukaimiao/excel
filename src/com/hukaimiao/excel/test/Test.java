/**  
 * @Title: Test.java
 * @Package com.hukaimiao.excel.test
 * @Description: TODO(用一句话描述该文件做什么)
 * @author hukaimiao  
 * @date 2014-9-12 上午8:49:09
 * @version V1.0  
 */
package com.hukaimiao.excel.test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;

import com.hukaimiao.excel.domain.Student;
import com.hukaimiao.excel.util.ExportExcel;
import com.hukaimiao.excel.util.InputExcel;
import com.hukaimiao.excel.util.POIExcelUtil;

/**
 * @ClassName: Test
 * @Description: TODO(这里用一句话描述这个类的作用)
 * @author hukaimiao
 * @date 2014-9-12 上午8:49:09
 */
public class Test {

	/*
	 * 导入
	 */
	@org.junit.Test
	public void inputExcel() {

		/**
		 * Workbook wb = POIExcelUtil.readExcel(excelFile);
		Map<String, Object> map = POIExcelUtil.readCell(wb.getSheetAt(0));
		for (String key : map.keySet()) {
			System.out.println(key + "-" + map.get(key));
		}
		 */
		Workbook wb = POIExcelUtil
				.readExcel("C:\\Users\\user\\Desktop\\未匹配的产品清单.xls");
		Map<String, Object> map = POIExcelUtil.readCell(wb.getSheetAt(0));
		for (String key : map.keySet()) {
			System.out.println(key + "-" + map.get(key));
		}
	}

	@org.junit.Test
	public void inputExcel2() {
		List<List<String>>  list = new ArrayList<List<String>>();
		String path = "C:\\Users\\user\\Desktop\\未匹配的产品清单.xls";// 文件路径
		try {
			File files = new File(path);
			String[][] result = InputExcel.getData(files, 1);
			if (result != null) {
				
				int rowLength = result.length;
				for (int i = 0; i < rowLength; i++) {
					
					 List<String> listSub=new ArrayList<String>();
					for (int j = 0; j < result[i].length; j++) {
						//System.out.println(result[i][j] + "单元格ID:" + i + " "
						//		+ j);
						 listSub.add(result[i][j]);
						 
						
					}
					 list.add(listSub);
				}
				
				System.out.println(list.toString());
				
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	@org.junit.Test
	public void testoutput() throws IOException {
		ExportExcel<Student> ex = new ExportExcel<Student>();
		String[] headers = { "学号", "姓名", "年龄", "性别", "出生日期" };
		List<Student> dataset = new ArrayList<Student>();
		dataset.add(new Student(10000001, "张三", 20, "男", new Date()));
		dataset.add(new Student(20000002, "李四", 24, "女", new Date()));
		dataset.add(new Student(30000003, "王五", 22, "女", new Date()));

		 OutputStream out = new FileOutputStream("d://a.xls");
		ex.exportExcel("测试", headers, dataset, out);
		out.close();
	}
	@org.junit.Test
	public void test(){
		Integer.parseInt("");
	}
}
