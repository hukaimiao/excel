/**  
 * @Title: InputExcel.java
 * @Package com.hukaimiao.excel.util
 * @Description: TODO(用一句话描述该文件做什么)
 * @author hukaimiao  
 * @date 2014-9-12 上午9:33:34
 * @version V1.0  
 */
package com.hukaimiao.excel.util;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * @ClassName: InputExcel
 * @Description: TODO(导入excel)
 * @author hukaimiao
 * @date 2014-9-12 上午9:33:34
 */
public class InputExcel {

	/**
	 * 读取EXCEL*
	 * 
	 * @param firstrow
	 *            从第几行开始读取
	 * @return 读取后返回数组
	 */
	@SuppressWarnings("deprecation")
	public static String[][] getData(File file, int firstrow)
			throws FileNotFoundException, IOException {
		List<String[]> result = new ArrayList<String[]>();
		int rowSize = 0;
		BufferedInputStream in = new BufferedInputStream(new FileInputStream(
				file));
		POIFSFileSystem fs = new POIFSFileSystem(in);
		HSSFWorkbook wb = new HSSFWorkbook(fs);
		HSSFCell cell = null;
		for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) {
			HSSFSheet st = wb.getSheetAt(sheetIndex);
			for (int rowIndex = firstrow; rowIndex <= st.getLastRowNum(); rowIndex++) {
				HSSFRow row = st.getRow(rowIndex);
				if (row == null) {
					continue;
				}
				int tempRowSize = row.getLastCellNum() + 1;
				if (tempRowSize > rowSize) {
					rowSize = tempRowSize;
				}
				String[] values = new String[rowSize];
				Arrays.fill(values, "");
				boolean hasValue = false;
				for (short columnIndex = 0; columnIndex <= row.getLastCellNum(); columnIndex++) {
					String value = "";
					cell = row.getCell(columnIndex);
					if (cell != null && !"".equals(cell)) {
						switch (cell.getCellType()) {
						case HSSFCell.CELL_TYPE_STRING:// 读取的格式为字符串
							value = cell.getStringCellValue();
							break;
						case HSSFCell.CELL_TYPE_NUMERIC:// 读取的格式为数组
							// 如果格式为日期格式，自定义格式输出
							if (HSSFDateUtil.isCellDateFormatted(cell)) {
								Date date = cell.getDateCellValue();
								if (date != null) {
									value = new SimpleDateFormat("yyyy-MM-dd")
											.format(date);
								} else {
									value = "";
								}
							} else {
								// 如果格式为数值，自定义格式输出
								//value = new DecimalFormat("0.00").format(cell
								//		.getNumericCellValue());
								value=String.valueOf((int)cell.getNumericCellValue());
							}
							break;
						case HSSFCell.CELL_TYPE_FORMULA:
							// 导入时如果为公式生成的数据则无值
							value = "";
							break; // 导入时如果为空
						case HSSFCell.CELL_TYPE_BLANK:
							break;
						case HSSFCell.CELL_TYPE_ERROR:
							value = "";
							break;

						// 导入时如果为BOOLEAN型 自定义格式输出
						case HSSFCell.CELL_TYPE_BOOLEAN:
							value = (cell.getBooleanCellValue() == true ? "Y"
									: "N");
							break;
						default:
							value = "";
						}
					}
					values[columnIndex] = rightTrim(value);
					hasValue = true;
				}
				if (hasValue) {
					result.add(values);
				}
			}
			in.close();
			String[][] returnArray = new String[result.size()][rowSize];
			for (int i = 0; i < returnArray.length; i++) {
				returnArray[i] = (String[]) result.get(i);
			}
			return returnArray;
		}
		return null;
	}

	public static String rightTrim(String str) {
		if (str == null) {
			return "";
		}
		int length = str.length();
		for (int i = length - 1; i >= 0; i--) {
			if (str.charAt(i) != 0x20) {
				break;
			}
			length--;
		}
		return str.substring(0, length);
	}
}
