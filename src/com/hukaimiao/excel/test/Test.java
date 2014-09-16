/**  
 * @Title: Test.java
 * @Package com.hukaimiao.excel.test
 * @Description: TODO(��һ�仰�������ļ���ʲô)
 * @author hukaimiao  
 * @date 2014-9-12 ����8:49:09
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
 * @Description: TODO(������һ�仰��������������)
 * @author hukaimiao
 * @date 2014-9-12 ����8:49:09
 */
public class Test {

	/*
	 * ����
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
				.readExcel("C:\\Users\\user\\Desktop\\δƥ��Ĳ�Ʒ�嵥.xls");
		Map<String, Object> map = POIExcelUtil.readCell(wb.getSheetAt(0));
		for (String key : map.keySet()) {
			System.out.println(key + "-" + map.get(key));
		}
	}

	@org.junit.Test
	public void inputExcel2() {
		List<List<String>>  list = new ArrayList<List<String>>();
		String path = "C:\\Users\\user\\Desktop\\δƥ��Ĳ�Ʒ�嵥.xls";// �ļ�·��
		try {
			File files = new File(path);
			String[][] result = InputExcel.getData(files, 1);
			if (result != null) {
				
				int rowLength = result.length;
				for (int i = 0; i < rowLength; i++) {
					
					 List<String> listSub=new ArrayList<String>();
					for (int j = 0; j < result[i].length; j++) {
						//System.out.println(result[i][j] + "��Ԫ��ID:" + i + " "
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
		String[] headers = { "ѧ��", "����", "����", "�Ա�", "��������" };
		List<Student> dataset = new ArrayList<Student>();
		dataset.add(new Student(10000001, "����", 20, "��", new Date()));
		dataset.add(new Student(20000002, "����", 24, "Ů", new Date()));
		dataset.add(new Student(30000003, "����", 22, "Ů", new Date()));

		 OutputStream out = new FileOutputStream("d://a.xls");
		ex.exportExcel("����", headers, dataset, out);
		out.close();
	}
	@org.junit.Test
	public void test(){
		Integer.parseInt("");
	}
}
