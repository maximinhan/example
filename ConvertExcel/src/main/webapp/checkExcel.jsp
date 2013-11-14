<%@ page language="java" contentType="text/html; charset=UTF-8"
	pageEncoding="UTF-8"%>
<%@ page import="java.io.*, java.util.*" %>
<%@ page
	import="com.oreilly.servlet.MultipartRequest,
                   com.oreilly.servlet.multipart.DefaultFileRenamePolicy,
                   java.util.*"%>
<%
	request.setCharacterEncoding("UTF-8");
	String savePath = getServletContext().getRealPath("/") + "/TempData/"; // 저장할 디렉토리 (절대경로)
	//convert excelconvert = new convert();
	int sizeLimit = 5 * 1024 * 1024; // 5메가까지 제한 넘어서면 예외발생

	File f = new File(savePath);
	if (!f.exists())//폴더가 없으면
	{
		f.mkdir();//폴더를 생성
	}

	try {

		MultipartRequest multi = new MultipartRequest(request,
				savePath, sizeLimit, "UTF-8",
				new DefaultFileRenamePolicy());
		Enumeration<?> fileNames = multi.getFileNames(); // file object의 이름 반환
		String fileName = multi.getFilesystemName(fileNames
				.nextElement().toString()); // 파일의 이름 얻기

		if (fileName == null) { // 파일이 업로드 되지 않았을때
			out.print("파일 업로드 실패");
		} else {
			String type = fileName.substring(fileName.length() - 1,
					fileName.length());

			if (type.equals("x")) {
				String from = savePath + fileName;
				String to = savePath + fileName.substring(0,
						fileName.length() - 1);

				request.setAttribute("from", from);
				request.setAttribute("to", to);

				RequestDispatcher dispatcher = request
						.getRequestDispatcher("./convert_xlsx.jsp");
				dispatcher.forward(request, response);

			} else if (type.equals("s")) {
				String from = savePath + fileName;
				String to = savePath + fileName + "x";

				request.setAttribute("from", from);
				request.setAttribute("to", to);

				RequestDispatcher dispatcher = request
						.getRequestDispatcher("./convert_xls.jsp");
				dispatcher.forward(request, response);
			}
		}

	} catch (Exception e) {
		e.printStackTrace();
	}
%>