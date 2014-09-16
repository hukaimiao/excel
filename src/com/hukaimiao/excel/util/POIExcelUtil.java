/**  
 * @Title: POIExcelUtil.java
 * @Package com.hukaimiao.excel.util
 * @Description: TODO(用一句话描述该文件做什么)
 * @author hukaimiao  
 * @date 2014-9-12 上午8:45:20
 * @version V1.0  
 */
package com.hukaimiao.excel.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.io.IOUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @ClassName: POIExcelUtil
 * @Description: TODO(POI操作类)
 * @author hukaimiao
 * @date 2014-9-12 上午8:45:20
 */
public class POIExcelUtil {

	// ------------------------写Excel-----------------------------------
	/**
	 * 创建workBook对象 xlsx(2007以上版本)
	 * 
	 * @return
	 */
	public static Workbook createWorkbook() {
		return createWorkbook(true);
	}

	/**
	 * 创建WorkBook对象
	 * 
	 * @param flag
	 *            true:xlsx(1997-2007) false:xls(2007以上)
	 * @return
	 */
	public static Workbook createWorkbook(boolean flag) {
		Workbook wb;
		if (flag) {
			wb = new XSSFWorkbook();
		} else {
			wb = new HSSFWorkbook();
		}
		return wb;
	}

	/**
	 * 添加图片
	 * 
	 * @param wb
	 *            workBook对象
	 * @param sheet
	 *            sheet对象
	 * @param picFileName
	 *            图片文件名称（全路径）
	 * @param picType
	 *            图片类型
	 * @param row
	 *            图片所在的行
	 * @param col
	 *            图片所在的列
	 */
	public static void addPicture(Workbook wb, Sheet sheet, String picFileName,
			int picType, int row, int col) {
		InputStream is = null;
		try {
			// 读取图片
			is = new FileInputStream(picFileName);
			byte[] bytes = IOUtils.toByteArray(is);
			int pictureIdx = wb.addPicture(bytes, picType);
			is.close();
			// 写图片
			CreationHelper helper = wb.getCreationHelper();
			Drawing drawing = sheet.createDrawingPatriarch();
			ClientAnchor anchor = helper.createClientAnchor();
			// 设置图片的位置
			anchor.setCol1(col);
			anchor.setRow1(row);
			Picture pict = drawing.createPicture(anchor, pictureIdx);

			pict.resize();
		} catch (Exception e) {
			try {
				if (is != null) {
					is.close();
				}
			} catch (IOException e1) {
				e1.printStackTrace();
			}
			e.printStackTrace();
		}
	}

	/**
	 * 创建Cell 默认为水平和垂直方式都是居中
	 * 
	 * @param style
	 *            CellStyle对象
	 * @param row
	 *            Row对象
	 * @param column
	 *            单元格所在的列
	 * @return
	 */
	public static Cell createCell(CellStyle style, Row row, short column) {
		return createCell(style, row, column, XSSFCellStyle.ALIGN_CENTER,
				XSSFCellStyle.ALIGN_CENTER);
	}

	/**
	 * 创建Cell并设置水平和垂直方式
	 * 
	 * @param style
	 *            CellStyle对象
	 * @param row
	 *            Row对象
	 * @param column
	 *            单元格所在的列
	 * @param halign
	 *            水平对齐方式：XSSFCellStyle.VERTICAL_CENTER.
	 * @param valign
	 *            垂直对齐方式：XSSFCellStyle.ALIGN_LEFT
	 */
	public static Cell createCell(CellStyle style, Row row, short column,
			short halign, short valign) {
		Cell cell = row.createCell(column);
		setAlign(style, halign, valign);
		cell.setCellStyle(style);
		return cell;
	}

	/**
	 * 合并单元格
	 * 
	 * @param sheet
	 * @param firstRow
	 *            开始行
	 * @param lastRow
	 *            最后行
	 * @param firstCol
	 *            开始列
	 * @param lastCol
	 *            最后列
	 */
	public static void mergeCell(Sheet sheet, int firstRow, int lastRow,
			int firstCol, int lastCol) {
		sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol,
				lastCol));
	}

	// ---------------------------------设置样式-----------------------

	/**
	 * 设置单元格对齐方式
	 * 
	 * @param style
	 * @param halign
	 * @param valign
	 * @return
	 */
	public static CellStyle setAlign(CellStyle style, short halign, short valign) {
		style.setAlignment(halign);
		style.setVerticalAlignment(valign);
		return style;
	}

	/**
	 * 设置单元格边框(四个方向的颜色一样)
	 * 
	 * @param style
	 *            style对象
	 * @param borderStyle
	 *            边框类型 ：dished-虚线 thick-加粗 double-双重 dotted-有点的
	 *            CellStyle.BORDER_THICK
	 * @param borderColor
	 *            颜色 IndexedColors.GREEN.getIndex()
	 * @return
	 */
	public static CellStyle setBorder(CellStyle style, short borderStyle,
			short borderColor) {

		// 设置底部格式（样式+颜色）
		style.setBorderBottom(borderStyle);
		style.setBottomBorderColor(borderColor);
		// 设置左边格式
		style.setBorderLeft(borderStyle);
		style.setLeftBorderColor(borderColor);
		// 设置右边格式
		style.setBorderRight(borderStyle);
		style.setRightBorderColor(borderColor);
		// 设置顶部格式
		style.setBorderTop(borderStyle);
		style.setTopBorderColor(borderColor);

		return style;
	}

	/**
	 * 自定义颜色（xssf)
	 * 
	 * @param style
	 *            xssfStyle
	 * @param red
	 *            RGB red (0-255)
	 * @param green
	 *            RGB green (0-255)
	 * @param blue
	 *            RGB blue (0-255)
	 */
	public static CellStyle setBackColorByCustom(XSSFCellStyle style, int red,
			int green, int blue) {
		// 设置前端颜色
		style.setFillForegroundColor(new XSSFColor(new java.awt.Color(red,
				green, blue)));
		// 设置填充模式
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);

		return style;
	}

	/**
	 * 设置前景颜色
	 * 
	 * @param style
	 *            style对象
	 * @param color
	 *            ：IndexedColors.YELLOW.getIndex()
	 * @return
	 */
	public static CellStyle setBackColor(CellStyle style, short color) {

		// 设置前端颜色
		style.setFillForegroundColor(color);
		// 设置填充模式
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);

		return style;
	}

	/**
	 * 设置背景颜色
	 * 
	 * @param style
	 *            style对象
	 * @param color
	 *            ：IndexedColors.YELLOW.getIndex()
	 * @param fillPattern
	 *            ：CellStyle.SPARSE_DOTS
	 * @return
	 */
	public static CellStyle setBackColor(CellStyle style, short backColor,
			short fillPattern) {

		// 设置背景颜色
		style.setFillBackgroundColor(backColor);

		// 设置填充模式
		style.setFillPattern(fillPattern);

		return style;
	}

	/**
	 * 
	 * 设置字体（简单的需求实现，如果复杂的字体，需要自己去实现）尽量重用
	 * 
	 * @param style
	 *            style对象
	 * @param fontSize
	 *            字体大小 shot(24)
	 * @param color
	 *            字体颜色 IndexedColors.YELLOW.getIndex()
	 * @param fontName
	 *            字体名称 "Courier New"
	 * @param
	 */
	public static CellStyle setFont(Font font, CellStyle style, short fontSize,
			short color, String fontName) {
		font.setFontHeightInPoints(color);
		font.setFontName(fontName);

		// font.setItalic(true);// 斜体
		// font.setStrikeout(true);//加干扰线

		font.setColor(color);// 设置颜色
		// Fonts are set into a style so create a new one to use.
		style.setFont(font);

		return style;

	}

	/**
	 * 
	 * @param createHelper
	 *            createHelper对象
	 * @param style
	 *            CellStyle对象
	 * @param formartData
	 *            date:"m/d/yy h:mm"; int:"#,###.0000" ,"0.0"
	 */
	public static CellStyle setDataFormat(CreationHelper createHelper,
			CellStyle style, String formartData) {

		style.setDataFormat(createHelper.createDataFormat().getFormat(
				formartData));

		return style;
	}

	/**
	 * 将Workbook写入文件
	 * 
	 * @param wb
	 *            workbook对象
	 * @param fileName
	 *            文件的全路径
	 * @return
	 */
	public static boolean createExcel(Workbook wb, String fileName) {
		boolean flag = true;
		FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream(fileName);
			wb.write(fileOut);
			fileOut.close();

		} catch (Exception e) {
			flag = false;
			if (fileOut != null) {
				try {
					fileOut.close();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
			}
			e.printStackTrace();
		}
		return flag;
	}

	// --------------------读取Excel-----------------------
	/**
	 * 读取Excel
	 * 
	 * @param filePathName
	 * @return
	 */
	public static Workbook readExcel(String filePathName) {
		InputStream inp = null;
		Workbook wb = null;
		try {
			inp = new FileInputStream(filePathName);
			wb = WorkbookFactory.create(inp);
			inp.close();

		} catch (Exception e) {
			try {
				if (null != inp) {
					inp.close();
				}
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			e.printStackTrace();
		}
		return wb;
	}

	/**
	 * 读取Cell的值
	 * 
	 * @param sheet
	 * @return
	 */
	public static Map readCell(Sheet sheet) {
		Map map = new HashMap();
		// 遍历所有行
		for (Row row : sheet) {
			// 便利所有列
			for (Cell cell : row) {
				// 获取单元格的类型
				CellReference cellRef = new CellReference(row.getRowNum(),
						cell.getColumnIndex());
				// System.out.print(cellRef.formatAsString());
				String key = cellRef.formatAsString();
				// System.out.print(" - ");

				switch (cell.getCellType()) {
				// 字符串
				case Cell.CELL_TYPE_STRING:
					map.put(key, cell.getRichStringCellValue().getString());
					// System.out.println(cell.getRichStringCellValue()
					// .getString());
					break;
				// 数字
				case Cell.CELL_TYPE_NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
						// System.out.println(cell.getDateCellValue());
						map.put(key, cell.getDateCellValue());
					} else {
						// System.out.println(cell.getNumericCellValue());
						map.put(key, cell.getNumericCellValue());
					}
					break;
				// boolean
				case Cell.CELL_TYPE_BOOLEAN:
					// System.out.println(cell.getBooleanCellValue());
					map.put(key, cell.getBooleanCellValue());
					break;
				// 方程式
				case Cell.CELL_TYPE_FORMULA:
					// System.out.println(cell.getCellFormula());
					map.put(key, cell.getCellFormula());
					break;
				// 空值
				default:
					System.out.println();
					map.put(key, "");
				}

			}
		}
		return map;

	}
}
