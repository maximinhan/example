<%@ page language="java" contentType="text/html; charset=UTF-8"
	pageEncoding="UTF-8"%>
<%@ page import="java.io.*, java.util.*"%>
<%@ page import="org.apache.poi.hssf.usermodel.*"%>
<%@ page import="org.apache.poi.ss.usermodel.*"%>
<%@ page import="org.apache.poi.xssf.usermodel.*"%>
<%@ page import="org.apache.poi.hssf.model.*"%>

<%
	String from = (String) request.getAttribute("from");
	String to = (String) request.getAttribute("to");
	FileInputStream fis = null;
	FileOutputStream fos = null;

	try {
		fis = new FileInputStream(from);
		Workbook workbook = WorkbookFactory.create(fis); //WorkbookFactory를 이용해 xls 워크북 생성       

		Workbook newWorkbook = new XSSFWorkbook(); //변환하여 내보낼 xlsx 워크북 생성
		Sheet sheet = workbook.getSheetAt(0);
		Sheet newSheet = newWorkbook.createSheet(); 

		//웹 화면에도 출력해 봅시다
		
		%><table border="1"><%

		Iterator<Row> rows = sheet.iterator();
		while (rows.hasNext()) {
			%><tr><%
			Row row = rows.next();
			Row newRow = newSheet.createRow(row.getRowNum());
	
			Iterator<Cell> cells = row.cellIterator();
			while (cells.hasNext()) {
				%><td><%
				Cell cell = cells.next();
				Cell newCell = newRow.createCell(cell.getColumnIndex());
				int type = cell.getCellType();
				
				switch (type) {
				case Cell.CELL_TYPE_BLANK:
					break;
				
				case Cell.CELL_TYPE_NUMERIC:					
					
					if (DateUtil.isCellDateFormatted(cell)) {
						out.println(cell.getDateCellValue());
						newCell.setCellValue(cell.getDateCellValue());
					} else {
						out.println(cell.getNumericCellValue());
						newCell.setCellValue(cell.getNumericCellValue());
					}
					break;
	
				case Cell.CELL_TYPE_STRING:
					out.print(cell.getStringCellValue() + ""); 
					newCell.setCellValue(cell.getStringCellValue());
					break;
	
				case Cell.CELL_TYPE_ERROR:
					out.print(cell.getErrorCellValue() + "");
					newCell.setCellErrorValue(cell.getErrorCellValue());
					break;
		
				case Cell.CELL_TYPE_BOOLEAN:
					out.print(cell.getBooleanCellValue() + "");
					newCell.setCellValue(cell.getBooleanCellValue());
					break;
	
				case Cell.CELL_TYPE_FORMULA:
					out.print(cell.getCellFormula() + "");
					newCell.setCellFormula(cell.getCellFormula());
					break;
				}
				%></td><%
			}			
			%></tr><%
		}
		
		fos = new FileOutputStream(to);
		newWorkbook.write(fos);	
		
	} catch (FileNotFoundException e) {
		e.printStackTrace();
	} catch (IOException e) {
		e.printStackTrace();
	} finally {
		fis.close();
		fos.close();
	}
%>
</table>
<bR>
<form method="post" action="./download.jsp">
	<input type="hidden" name="path" value="<%=to %>">
	<input type="submit" value="다운로드">
</form>