<%@ page language="java" contentType="text/html; charset=UTF-8"
	pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Convert Excel</title>
</head>
<body>
	<h1>Hello POI!</h1>
	.xlsx (엑셀 2007이상 버전) 은 .xls 로 .xls (엑셀 2007미만 버전) 은 .xlsx 로 한번 변환해 봅시다.<br><br>
	현 예제는 업로드 예제가 아니므로 파일 유효성 검사등은 귀찮으니 생략합시다.<br>

	<form action="checkExcel.jsp" enctype="multipart/form-data" method="post">
		엑셀 파일을 선택하세요 : <input type="file" name="excel" size="40"><br>
		<input type="submit" value="변환!">
	</form>
</body>
</html>