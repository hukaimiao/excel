/**  
 * @Title: OutputExcelAction.java
 * @Package com.hukaimiao.excel.action
 * @Description: TODO(用一句话描述该文件做什么)
 * @author hukaimiao  
 * @date 2014-9-11 下午4:16:55
 * @version V1.0  
 */
package com.hukaimiao.excel.action;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.struts2.ServletActionContext;

import com.hukaimiao.excel.domain.Student;
import com.hukaimiao.excel.util.ExcelWorkSheet;
import com.hukaimiao.excel.util.ExportExcel;
import com.hukaimiao.excel.util.InputExcel;
import com.hukaimiao.excel.util.POIExcelUtil;
import com.opensymphony.xwork2.ActionContext;
import com.opensymphony.xwork2.ActionSupport;

/**
 * @ClassName: OutputExcelAction
 * @Description: TODO(导出excel)
 * @author hukaimiao
 * @date 2014-9-11 下午4:16:55
 */
public class ManageExcelAction extends ActionSupport {

	/**
	 * @Fields serialVersionUID : TODO(用一句话描述这个变量表示什么)
	 */
	private static final long serialVersionUID = 1L;
	ActionContext cxt = ActionContext.getContext();
	HttpServletRequest request = (HttpServletRequest) cxt
			.get(ServletActionContext.HTTP_REQUEST);
	HttpServletResponse response = (HttpServletResponse) cxt
			.get(ServletActionContext.HTTP_RESPONSE);

	private List<Student> list;
	// 导出excel需要的变量
	private File excelFile; // 上传的文件

	private String excelFileFileName; // 保存原始文件名

	// 将Excel文件解析完毕后信息存放到这个对象中
	private ExcelWorkSheet<Student> excelWorkSheet;

	/**
	 * 导出excel
	 * 
	 * @throws IOException
	 */
	public String OutputExcel() throws IOException {
		// excel表格名字
		String filedisplay = "未匹配的产品清单.xls";
		// excel标题
		String title = "POI测试";
		filedisplay = URLEncoder.encode(filedisplay, "UTF-8");

		request.setCharacterEncoding("UTF-8");
		response.setCharacterEncoding("UTF-8");
		response.setContentType("application/x-download");
		response.addHeader("Content-Disposition", "attachment;filename="
				+ filedisplay);

		ExportExcel<Student> ex = new ExportExcel<Student>();
		String[] headers = { "学号", "姓名", "年龄", "性别", "出生日期" };
		List<Student> dataset = new ArrayList<Student>();
		dataset.add(new Student(10000001, "张三", 20, "男", new Date()));
		dataset.add(new Student(20000002, "李四", 24, "女", new Date()));
		dataset.add(new Student(30000003, "王五", 22, "女", new Date()));

		// OutputStream out = new FileOutputStream("d://a.xls");
		OutputStream out = response.getOutputStream();
		ex.exportExcel(title, headers, dataset, out);
		out.close();
		// JOptionPane.showMessageDialog(null, "导出成功!");
		return null;
	}

	/**
	 * 导入excel
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	public String InputExcel() throws FileNotFoundException, IOException {

		Student student = null;
		// String path = "C:\\Users\\user\\Desktop\\未匹配的产品清单.xls";// 文件路径
		try {
			// File files = new File(path);
			String[][] result = InputExcel.getData(excelFile, 1);

			if (result != null) {
				list = new ArrayList<Student>();

				int rowLength = result.length;
				Date date ;
				for (int i = 0; i < rowLength; i++) {
					SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
					// for (int j = 0; j < result[i].length; j++) {

					// System.out.println(result[i][j] + "单元格ID:" + i + " "+ j);
					student = new Student();
					if (result[i][0] != null && !"".equals(result[i][0])) {
						student.setId(Integer.parseInt(result[i][0]));
					}
					if (result[i][1] != null && !"".equals(result[i][1])) {
						student.setName(result[i][1]);
					}else{
						student.setName("null");
					}
					if (result[i][2] != null && !"".equals(result[i][2])) {
						student.setAge(Integer.parseInt(result[i][2]));
					}
					if (result[i][3] != null && !"".equals(result[i][3])) {
						student.setSex(result[i][3]);
					}else{
						student.setSex("null");
					}
					if (result[i][4] != null && !"".equals(result[i][4])) {
						date= sdf.parse(result[i][4]);
						student.setBirthday(date);
					}
					
					list.add(student);
					// }
				}

				System.out.println(list.size());
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

		return SUCCESS;
	}

	// 判断文件类型
	public Workbook createWorkBook(InputStream is) throws IOException {
		if (excelFileFileName.toLowerCase().endsWith("xls")) {
			return new HSSFWorkbook(is);
		}
		if (excelFileFileName.toLowerCase().endsWith("xlsx")) {
			return new XSSFWorkbook(is);
		}
		return null;
	}

	public File getExcelFile() {
		return excelFile;
	}

	public void setExcelFile(File excelFile) {
		this.excelFile = excelFile;
	}

	public String getExcelFileFileName() {
		return excelFileFileName;
	}

	public void setExcelFileFileName(String excelFileFileName) {
		this.excelFileFileName = excelFileFileName;
	}

	public ExcelWorkSheet<Student> getExcelWorkSheet() {
		return excelWorkSheet;
	}

	public void setExcelWorkSheet(ExcelWorkSheet<Student> excelWorkSheet) {
		this.excelWorkSheet = excelWorkSheet;
	}

	public List<Student> getList() {
		return list;
	}

	public void setList(List<Student> list) {
		this.list = list;
	}

}
