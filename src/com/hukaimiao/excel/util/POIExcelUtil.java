/**  
 * @Title: POIExcelUtil.java
 * @Package com.hukaimiao.excel.util
 * @Description: TODO(��һ�仰�������ļ���ʲô)
 * @author hukaimiao  
 * @date 2014-9-12 ����8:45:20
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
 * @Description: TODO(POI������)
 * @author hukaimiao
 * @date 2014-9-12 ����8:45:20
 */
public class POIExcelUtil {

	// ------------------------дExcel-----------------------------------
	/**
	 * ����workBook���� xlsx(2007���ϰ汾)
	 * 
	 * @return
	 */
	public static Workbook createWorkbook() {
		return createWorkbook(true);
	}

	/**
	 * ����WorkBook����
	 * 
	 * @param flag
	 *            true:xlsx(1997-2007) false:xls(2007����)
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
	 * ���ͼƬ
	 * 
	 * @param wb
	 *            workBook����
	 * @param sheet
	 *            sheet����
	 * @param picFileName
	 *            ͼƬ�ļ����ƣ�ȫ·����
	 * @param picType
	 *            ͼƬ����
	 * @param row
	 *            ͼƬ���ڵ���
	 * @param col
	 *            ͼƬ���ڵ���
	 */
	public static void addPicture(Workbook wb, Sheet sheet, String picFileName,
			int picType, int row, int col) {
		InputStream is = null;
		try {
			// ��ȡͼƬ
			is = new FileInputStream(picFileName);
			byte[] bytes = IOUtils.toByteArray(is);
			int pictureIdx = wb.addPicture(bytes, picType);
			is.close();
			// дͼƬ
			CreationHelper helper = wb.getCreationHelper();
			Drawing drawing = sheet.createDrawingPatriarch();
			ClientAnchor anchor = helper.createClientAnchor();
			// ����ͼƬ��λ��
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
	 * ����Cell Ĭ��Ϊˮƽ�ʹ�ֱ��ʽ���Ǿ���
	 * 
	 * @param style
	 *            CellStyle����
	 * @param row
	 *            Row����
	 * @param column
	 *            ��Ԫ�����ڵ���
	 * @return
	 */
	public static Cell createCell(CellStyle style, Row row, short column) {
		return createCell(style, row, column, XSSFCellStyle.ALIGN_CENTER,
				XSSFCellStyle.ALIGN_CENTER);
	}

	/**
	 * ����Cell������ˮƽ�ʹ�ֱ��ʽ
	 * 
	 * @param style
	 *            CellStyle����
	 * @param row
	 *            Row����
	 * @param column
	 *            ��Ԫ�����ڵ���
	 * @param halign
	 *            ˮƽ���뷽ʽ��XSSFCellStyle.VERTICAL_CENTER.
	 * @param valign
	 *            ��ֱ���뷽ʽ��XSSFCellStyle.ALIGN_LEFT
	 */
	public static Cell createCell(CellStyle style, Row row, short column,
			short halign, short valign) {
		Cell cell = row.createCell(column);
		setAlign(style, halign, valign);
		cell.setCellStyle(style);
		return cell;
	}

	/**
	 * �ϲ���Ԫ��
	 * 
	 * @param sheet
	 * @param firstRow
	 *            ��ʼ��
	 * @param lastRow
	 *            �����
	 * @param firstCol
	 *            ��ʼ��
	 * @param lastCol
	 *            �����
	 */
	public static void mergeCell(Sheet sheet, int firstRow, int lastRow,
			int firstCol, int lastCol) {
		sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol,
				lastCol));
	}

	// ---------------------------------������ʽ-----------------------

	/**
	 * ���õ�Ԫ����뷽ʽ
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
	 * ���õ�Ԫ��߿�(�ĸ��������ɫһ��)
	 * 
	 * @param style
	 *            style����
	 * @param borderStyle
	 *            �߿����� ��dished-���� thick-�Ӵ� double-˫�� dotted-�е��
	 *            CellStyle.BORDER_THICK
	 * @param borderColor
	 *            ��ɫ IndexedColors.GREEN.getIndex()
	 * @return
	 */
	public static CellStyle setBorder(CellStyle style, short borderStyle,
			short borderColor) {

		// ���õײ���ʽ����ʽ+��ɫ��
		style.setBorderBottom(borderStyle);
		style.setBottomBorderColor(borderColor);
		// ������߸�ʽ
		style.setBorderLeft(borderStyle);
		style.setLeftBorderColor(borderColor);
		// �����ұ߸�ʽ
		style.setBorderRight(borderStyle);
		style.setRightBorderColor(borderColor);
		// ���ö�����ʽ
		style.setBorderTop(borderStyle);
		style.setTopBorderColor(borderColor);

		return style;
	}

	/**
	 * �Զ�����ɫ��xssf)
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
		// ����ǰ����ɫ
		style.setFillForegroundColor(new XSSFColor(new java.awt.Color(red,
				green, blue)));
		// �������ģʽ
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);

		return style;
	}

	/**
	 * ����ǰ����ɫ
	 * 
	 * @param style
	 *            style����
	 * @param color
	 *            ��IndexedColors.YELLOW.getIndex()
	 * @return
	 */
	public static CellStyle setBackColor(CellStyle style, short color) {

		// ����ǰ����ɫ
		style.setFillForegroundColor(color);
		// �������ģʽ
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);

		return style;
	}

	/**
	 * ���ñ�����ɫ
	 * 
	 * @param style
	 *            style����
	 * @param color
	 *            ��IndexedColors.YELLOW.getIndex()
	 * @param fillPattern
	 *            ��CellStyle.SPARSE_DOTS
	 * @return
	 */
	public static CellStyle setBackColor(CellStyle style, short backColor,
			short fillPattern) {

		// ���ñ�����ɫ
		style.setFillBackgroundColor(backColor);

		// �������ģʽ
		style.setFillPattern(fillPattern);

		return style;
	}

	/**
	 * 
	 * �������壨�򵥵�����ʵ�֣�������ӵ����壬��Ҫ�Լ�ȥʵ�֣���������
	 * 
	 * @param style
	 *            style����
	 * @param fontSize
	 *            �����С shot(24)
	 * @param color
	 *            ������ɫ IndexedColors.YELLOW.getIndex()
	 * @param fontName
	 *            �������� "Courier New"
	 * @param
	 */
	public static CellStyle setFont(Font font, CellStyle style, short fontSize,
			short color, String fontName) {
		font.setFontHeightInPoints(color);
		font.setFontName(fontName);

		// font.setItalic(true);// б��
		// font.setStrikeout(true);//�Ӹ�����

		font.setColor(color);// ������ɫ
		// Fonts are set into a style so create a new one to use.
		style.setFont(font);

		return style;

	}

	/**
	 * 
	 * @param createHelper
	 *            createHelper����
	 * @param style
	 *            CellStyle����
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
	 * ��Workbookд���ļ�
	 * 
	 * @param wb
	 *            workbook����
	 * @param fileName
	 *            �ļ���ȫ·��
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

	// --------------------��ȡExcel-----------------------
	/**
	 * ��ȡExcel
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
	 * ��ȡCell��ֵ
	 * 
	 * @param sheet
	 * @return
	 */
	public static Map readCell(Sheet sheet) {
		Map map = new HashMap();
		// ����������
		for (Row row : sheet) {
			// ����������
			for (Cell cell : row) {
				// ��ȡ��Ԫ�������
				CellReference cellRef = new CellReference(row.getRowNum(),
						cell.getColumnIndex());
				// System.out.print(cellRef.formatAsString());
				String key = cellRef.formatAsString();
				// System.out.print(" - ");

				switch (cell.getCellType()) {
				// �ַ���
				case Cell.CELL_TYPE_STRING:
					map.put(key, cell.getRichStringCellValue().getString());
					// System.out.println(cell.getRichStringCellValue()
					// .getString());
					break;
				// ����
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
				// ����ʽ
				case Cell.CELL_TYPE_FORMULA:
					// System.out.println(cell.getCellFormula());
					map.put(key, cell.getCellFormula());
					break;
				// ��ֵ
				default:
					System.out.println();
					map.put(key, "");
				}

			}
		}
		return map;

	}
}
