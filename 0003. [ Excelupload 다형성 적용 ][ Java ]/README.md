# [ Excelupload에 다형성 적용 ][ Java ] <br><br>

- 설명 : <br>

     poi 라이브러리의 경우 엑셀 파일의 확장자에 따라 다른 클래스를 사용한다. <br>
     
     xls는 HSSF 관련 클래스들을 사용하고 <br>
     xlsx는 XSSF 관련 클래스들을 사용한다. <br>
     
     이 경우 사용하는 클래스에 따라 두가지 로직으로 나뉘게 되고 중복코드가 발생한다 <br>
     다형성을 사용하면 한 로직으로 처리할 수 있다. 
<br><br><br><br>





### 1. 인터페이스 참조변수를 선언한다. <br>

HSSF와 XSSF 관련 클래스들은 같은 인터페이스를 구현한 클래스들이다. <br>
인터페이스로 참조변수를 선언하면 인터페이스를 구현해서 만든 모든 클래스의 객체를 담을 수 있다. <br>
HSSF와 XSSF 관련 클래스들의 인터페이스로 참조변수를 선언한다. <br>

```java
Workbook workbook = null;
Sheet sheet = null;
Row row = null;
```
<br><br><br><br>





### 2. 엑셀파일의 확장자에 맞는 객체들을 생성한다. <br>

 확장자에 맞게 workbook클래스를 생성하고 <br>
 workbook클래스에서 sheet를 가져와서 확장자에 맞는 sheet클래스로 업캐스팅 해준다. <br>
 sheet클래스에서 row클래스는 그냥 가져와서 사용하면 된다. <br>


```java
if(extension.equals("xls")) {	
	workbook = new HSSFWorkbook(fis); 
	sheet = (HSSFSheet)workbook.getSheetAt(0); 
}else {						
	workbook = new XSSFWorkbook(fis); 
	sheet = (XSSFSheet)workbook.getSheetAt(0);  
};

row = sheet.getRow(0); 
```
<br><br><br><br>





### 결과  <br>

- 중복 로직 제거<br>






























<br><br><br><br>




```java
 
```









