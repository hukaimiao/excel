package com.hukaimiao.excel.servlet;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileItemFactory;
import org.apache.commons.fileupload.FileUploadBase.FileSizeLimitExceededException;
import org.apache.commons.fileupload.FileUploadException;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;

import com.hukaimiao.excel.domain.Student;
import com.hukaimiao.excel.util.InputExcel;

public class InputExcelServlet extends HttpServlet {

	/**
	 * @Fields serialVersionUID : TODO(用一句话描述这个变量表示什么)
	 */
	private static final long serialVersionUID = 1L;

	private List<Student> list;

	public InputExcelServlet() {
		super();
	}

	public void destroy() {
		super.destroy();
	}

	public void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		Student student = null;
		request.setCharacterEncoding("UTF-8");
		response.setContentType("text/html; charset=UTF-8");

		// 保存路径
		String savePath = getServletContext().getRealPath("/upload");
		File saveDir = new File(savePath);
		// 如果目录不存在，就创建目录
		if (!saveDir.exists()) {
			saveDir.mkdir();
		}

		// 创建文件上传核心类
		DiskFileItemFactory factory = new DiskFileItemFactory();
		ServletFileUpload sfu = new ServletFileUpload(factory);
		// 设置编码
		sfu.setHeaderEncoding("UTF-8");
		// 设置上传的单个文件的最大字节数为2M
		sfu.setFileSizeMax(1024 * 1024 * 200);
		// 设置整个表单的最大字节数为10M
		sfu.setSizeMax(1024 * 1024 * 100);

		try {
			// 处理表单请求
			List<FileItem> itemList = sfu.parseRequest(request);
			for (FileItem fileItem : itemList) {
				// 对应表单中的控件的name
				String fieldName = fileItem.getFieldName();
				System.out.println("控件名称：" + fieldName);
				// 如果是普通表单控件
				if (fileItem.isFormField()) {
					String value = fileItem.getString();
					// 重新编码,解决乱码
					value = new String(value.getBytes("ISO-8859-1"), "UTF-8");
					System.out.println("普通内容：" + fieldName + "=" + value);
					// 上传文件
				} else {
					// 获得文件大小
					Long size = fileItem.getSize();
					// 获得文件名
					String fileName = fileItem.getName();

					// 设置不允许上传的文件格式
					if (fileName.endsWith(".xls") ||fileName.endsWith(".xlsx")) {
						// 将文件保存到指定的路径
						File excelFile = new File(savePath, fileName);
						 fileItem.write(excelFile);
						String[][] result = InputExcel.getData(excelFile, 1);
						if (result != null) {
							File delPhotoPath  = new File(savePath+"\\"+fileName);
							System.out.println(savePath+fileName);
							if(delPhotoPath.exists()){
							  delPhotoPath.delete();
							}
							list = new ArrayList<Student>();
							int rowLength = result.length;
							Date date;
							for (int i = 0; i < rowLength; i++) {
								SimpleDateFormat sdf = new SimpleDateFormat(
										"yyyy-MM-dd");
								student = new Student();
								if (result[i][0] != null
										&& !"".equals(result[i][0])) {
									student.setId(Integer
											.parseInt(result[i][0]));
								}
								if (result[i][1] != null
										&& !"".equals(result[i][1])) {
									student.setName(result[i][1]);
								} else {
									student.setName("null");
								}
								if (result[i][2] != null
										&& !"".equals(result[i][2])) {
									student.setAge(Integer
											.parseInt(result[i][2]));
								}
								if (result[i][3] != null
										&& !"".equals(result[i][3])) {
									student.setSex(result[i][3]);
								} else {
									student.setSex("null");
								}
								if (result[i][4] != null
										&& !"".equals(result[i][4])) {
									date = sdf.parse(result[i][4]);
									student.setBirthday(date);
								}

								list.add(student);
							}
							request.setAttribute("list", list);
							request.getRequestDispatcher("/inputdata.jsp")
									.forward(request, response);
						}
						
					} else {
						request.setAttribute("msg", "不允许上传的类型！");
					}
				}
			}
		} catch (FileSizeLimitExceededException e) {
			request.setAttribute("msg", "文件太大");
		} catch (FileUploadException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {

		this.doGet(request, response);
	}

	public void init() throws ServletException {

	}

	// 统计结果用，采用Character即char做键（Key）
	private Map<Character, Integer> countMap = new HashMap<Character, Integer>();

	public void countChar(String str) {
		char[] chars = str.toCharArray();// 将字符串转换成字符char数组
		// 循环，开始统计
		for (char ch : chars) {
			// 判断字符是否存在
			if (!countMap.containsKey(ch)) {
				// 不存在，在Map中加一个，并设置初始值为0
				countMap.put(ch, 0);
			}
			// 计数，将值+1
			int count = countMap.get(ch);
			countMap.put(ch, count + 1);
		}

		// 输出结果
		Set<Character> keys = countMap.keySet();
		for (Character ch : keys) {
			System.out.println("字符" + ch + "出现次数：" + countMap.get(ch));
		}

	}

	public static void main(String[] args) {
		// 测试方法
		InputExcelServlet test = new InputExcelServlet();
		test.countChar("11Adfasadfadaere"); // 注：不支持中文
	}

}
