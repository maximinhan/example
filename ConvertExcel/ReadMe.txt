src>main>webapp 에 예제코드가 작성되어 있습니다.
모두 jsp로만 작성하였습니다. (Class 미사용)

HSSFWorkbook 은 엑셀 97~2003 파일 핸들링에 대응하며 
XSSFWorkbook 은 엑셀 2007~  파일 핸들링에 대응하는 클래스 입니다.


현재 날짜 관련 포맷을 읽는 것에 문제가 있습니다.
날짜형식의 셀임을 인지는 하나 날짜 형식까지는 완전히 읽지 못하여 정수형 숫자로 읽습니다.
이부분은 cell의 스타일 등을 이용하여 해결해야 하는 것으로 보이나 현재는 구현하지 않았습니다.