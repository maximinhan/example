<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
 <%@ page import="java.util.*, java.io.*" %>
<%@ page import="org.apache.poi.hssf.usermodel.*"%>
<%@ page import="org.apache.poi.ss.usermodel.*"%>
<%@ page import="org.apache.poi.xssf.usermodel.*"%>
<%@ page import="org.apache.poi.hssf.model.*"%>
<%
String path = (String) request.getParameter("path");

//웹 상에서의 파일 다운로드를 구현하기 위해 헤더를 추가
response.setHeader("Content-Disposition", "attachment; filename="+new String(path.getBytes("UTF-8"),"UTF-8"));

FileInputStream fis = new FileInputStream(path);
Workbook workbook = WorkbookFactory.create(fis);
	
out.clear();
out=pageContext.pushBody();
//Jsp 내장 out객체를 초기화 해준 후 새로운 out객체를 얻어옵니다.
ServletOutputStream fileout = response.getOutputStream();

workbook.write(fileout);
fis.close();

//현 예제에선 논의 대상이 아닌 예외처리는 패스.
%>