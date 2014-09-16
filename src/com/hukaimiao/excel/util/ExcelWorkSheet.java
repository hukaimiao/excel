/**  
 * @Title: ExcelWorkSheet.java
 * @Package com.hukaimiao.excel.util
 * @Description: TODO(��һ�仰�������ļ���ʲô)
 * @author hukaimiao  
 * @date 2014-9-12 ����9:02:51
 * @version V1.0  
 */
package com.hukaimiao.excel.util;

import java.util.ArrayList;
import java.util.List;

/**
 * @ClassName: ExcelWorkSheet
 * @Description: TODO(������һ�仰��������������)
 * @author hukaimiao
 * @date 2014-9-12 ����9:02:51
 */
public class ExcelWorkSheet<T> {
	private String sheetName;
	private List<T> data = new ArrayList<T>(); // ������
	private List<String> columns; // ����

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public List<T> getData() {
		return data;
	}

	public void setData(List<T> data) {
		this.data = data;
	}

	public List<String> getColumns() {
		return columns;
	}

	public void setColumns(List<String> columns) {
		this.columns = columns;
	}
}