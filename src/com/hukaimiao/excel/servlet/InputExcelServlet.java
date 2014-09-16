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
	 * @Fields serialVersionUID : TODO(��һ�仰�������������ʾʲô)
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

		// ����·��
		String savePath = getServletContext().getRealPath("/upload");
		File saveDir = new File(savePath);
		// ���Ŀ¼�����ڣ��ʹ���Ŀ¼
		if (!saveDir.exists()) {
			saveDir.mkdir();
		}

		// �����ļ��ϴ�������
		DiskFileItemFactory factory = new DiskFileItemFactory();
		ServletFileUpload sfu = new ServletFileUpload(factory);
		// ���ñ���
		sfu.setHeaderEncoding("UTF-8");
		// �����ϴ��ĵ����ļ�������ֽ���Ϊ2M
		sfu.setFileSizeMax(1024 * 1024 * 200);
		// ����������������ֽ���Ϊ10M
		sfu.setSizeMax(1024 * 1024 * 100);

		try {
			// ���������
			List<FileItem> itemList = sfu.parseRequest(request);
			for (FileItem fileItem : itemList) {
				// ��Ӧ���еĿؼ���name
				String fieldName = fileItem.getFieldName();
				System.out.println("�ؼ����ƣ�" + fieldName);
				// �������ͨ���ؼ�
				if (fileItem.isFormField()) {
					String value = fileItem.getString();
					// ���±���,�������
					value = new String(value.getBytes("ISO-8859-1"), "UTF-8");
					System.out.println("��ͨ���ݣ�" + fieldName + "=" + value);
					// �ϴ��ļ�
				} else {
					// ����ļ���С
					Long size = fileItem.getSize();
					// ����ļ���
					String fileName = fileItem.getName();

					// ���ò������ϴ����ļ���ʽ
					if (fileName.endsWith(".xls") ||fileName.endsWith(".xlsx")) {
						// ���ļ����浽ָ����·��
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
						request.setAttribute("msg", "�������ϴ������ͣ�");
					}
				}
			}
		} catch (FileSizeLimitExceededException e) {
			request.setAttribute("msg", "�ļ�̫��");
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

	// ͳ�ƽ���ã�����Character��char������Key��
	private Map<Character, Integer> countMap = new HashMap<Character, Integer>();

	public void countChar(String str) {
		char[] chars = str.toCharArray();// ���ַ���ת�����ַ�char����
		// ѭ������ʼͳ��
		for (char ch : chars) {
			// �ж��ַ��Ƿ����
			if (!countMap.containsKey(ch)) {
				// �����ڣ���Map�м�һ���������ó�ʼֵΪ0
				countMap.put(ch, 0);
			}
			// ��������ֵ+1
			int count = countMap.get(ch);
			countMap.put(ch, count + 1);
		}

		// ������
		Set<Character> keys = countMap.keySet();
		for (Character ch : keys) {
			System.out.println("�ַ�" + ch + "���ִ�����" + countMap.get(ch));
		}

	}

	public static void main(String[] args) {
		// ���Է���
		InputExcelServlet test = new InputExcelServlet();
		test.countChar("11Adfasadfadaere"); // ע����֧������
	}

}
