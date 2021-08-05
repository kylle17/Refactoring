# [ Excel Download ][ Delphi ] <br>

- 설명 : 그리드에 있는 하기휴가 계획서 내용을 엑셀 파일로 출력하는 기능을 리팩토링 <br><br><br><br>





### 1. 클래스 하나를 모듈 단위로 정함 <br>

Delphi의 경우 클래스를 pas파일 상단에 먼저 선언해줘야 한다. <br>
상단 선언부에 들어가는 내용은 필드 , 생성자 , 프로시저 , 함수이다. <br>

- 상단 선언부
```delphi
  type
    Excel_module_class = class
    public
       dept_name :String;
       excel :Variant;
       grid1 , grid2 :TAdvStringGrid;
       grid1_data_arr , grid2_data_arr :OleVariant;
       grid1_arr_row_length , grid1_arr_col_length : integer;
       grid2_arr_row_length , grid2_arr_col_length : integer;
       excel_start_cell_1 , excel_end_cell_1  :Variant;
       excel_start_cell_2 , excel_end_cell_2  :Variant;

       constructor Create( temp_dept_name :String ; temp_grid1 , temp_grid2 :TAdvStringGrid );
       procedure grid_is_empty();
       function make_excel() :Variant ;
       procedure set_excel_print_page();
       procedure set_field_grid_data_arr();
       procedure set_field_grid1_arr_length();
       procedure set_field_excel_start_end_cell();
       procedure set_excel_sheet_name( sheet_name:String );
       procedure set_cexcel_title( title:String );
       procedure set_title_design( row :integer );
       function make_grid_data_arr( grid:TAdvStringGrid ) :OleVariant;
       function get_grid_arr_row_length( grid_data_arr :OleVariant ) :integer;
       function get_grid_arr_col_length( grid_data_arr :OleVariant ) :integer;
       procedure insert_data_to_excel();
       procedure set_aligment_center();
       procedure set_height_width();
       procedure set_excel_line();
       procedure set_color();
       procedure set_font();

  end;// Excel_module_class 끝
```
<br><br><br><br>





### 2. 엑셀 출력 작업 전체흐름 <br>

 각 작업을 함수단위로 분리해서 필요한 부분만 찾아서 수정할 수 있고 <br>
 로직이 분리되어 있기 때문에 수정시에 다른 로직에 영향을 미치지 않는다. <br>
  
```delphi
constructor THB40010.Excel_module_class.Create( temp_dept_name :String ; temp_grid1 , temp_grid2 :TAdvStringGrid );
var
   sheet_name , title :String;
begin
  inherited Create;

  // 필드 값으로 저장.
  dept_name := temp_dept_name;
  grid1 := temp_grid1;
  grid2 := temp_grid2;

  // 그리드에 데이터가 없는 경우 처리
  grid_is_empty();

  // 엑셀 객체 생성.
  excel := make_excel();

  // 엑셀 페이지 설정
  set_excel_print_page();

  // 그리드의 데이터를 배열로 변
  set_field_grid_data_arr();
  set_field_grid1_arr_length();
  
  // 시작지점 설정
  set_field_excel_start_end_cell();

  // 시트 이름 설정
  sheet_name := '하기휴가 계획서';
  set_excel_sheet_name( sheet_name );

  // 엑셀에 데이터 출력
  insert_data_to_excel();
  
  // Cell 정렬 설정
  set_aligment_center();
  
  // Cell 높이,넓이 설정
  set_height_width();
  
  // Cell 테두리 설정
  set_excel_line();
  
  // Cell 색 설정
  set_color();
  
  // Cell 글씨체 설정
  set_font();

  // 제목 설정
  title := dept_name + '  하기휴가 계획서';
  set_cexcel_title( title);

  excel.Visible:= True;
end;
```
<br><br><br><br>





### 3. 그리드 데이터를 배열로 변환해서 필드에 저장. <br>

delphi의 엑셀 컴포넌트는 배열자료형을 사용하기 때문에 그리드 데이터를 배열로 변환한다. <br>
그렇게 변환한 배열의 경우 모듈 내에서 여러번 사용하기 때문에 필드에 저장해서 사용한다. <br>
필드값으로 구성해 두면 매번 인자값으로 전달하지 않을 수 있다. <br>
단점은 필드값의 선언부가 상단에 있기 때문에 함수에서 코드를 볼 때 왔다갔다 해야 하는 경우가 발생한다. <br>
이부분은 최대한 naming을 잘해서 상단을 찾아보지 않고 바로 이해할 수 있게 한다. <br>

- 그리드 데이터를 필드에 저장.
```delphi
constructor THB40010.Excel_module_class.Create( temp_dept_name :String ; temp_grid1 , temp_grid2 :TAdvStringGrid );
var
   sheet_name , title :String;
begin
  생략...

  grid1 := temp_grid1;
  grid2 := temp_grid2;  
  
  생략...
end;  
```
<br>

- 그리드 데이터를 배열로 변환.
```delphi
function THB40010.Excel_module_class.make_grid_data_arr( grid:TAdvStringGrid ) :OleVariant;
var
  grid_data_arr:OleVariant;
  row , col : integer;
begin
   grid_data_arr := VarArrayCreate( [ 0 , grid.RowCount , 0 , grid.ColCount ] , VarVariant );

   For row := 0 to grid.RowCount do
   begin
      For col := 0 to grid.ColCount do
      begin
          grid_data_arr[ row , col ] := '' + grid.cells[ col , row];
      end;
   end;
   result := grid_data_arr;
end;
```
<br>

- 변환된 데이터를 필드에 저장.
```delphi
procedure THB40010.Excel_module_class.set_field_grid_data_arr();
begin
     grid1_data_arr := make_grid_data_arr( grid1 );
     grid2_data_arr := make_grid_data_arr( grid2 );
end;
```
<br>

- 배열의 길이도 자주 사용하기 때문에 필드로 구성.
```delphi
procedure THB40010.Excel_module_class.set_field_grid1_arr_length();
begin
  grid1_arr_row_length := get_grid_arr_row_length( grid1_data_arr );
  grid1_arr_col_length := get_grid_arr_col_length( grid1_data_arr );
  grid2_arr_row_length := get_grid_arr_row_length( grid2_data_arr );
  grid2_arr_col_length := get_grid_arr_col_length( grid2_data_arr );
end;
```
<br><br><br><br>





### 4. 그리드에 데이터가 없는 경우 처리. <br>

그리드에 데이터가 없으면 Exception을 발생시킨다. 
```delphi
procedure THB40010.Excel_module_class.grid_is_empty();
begin
  if grid1.Cells[0, grid1.FixedRows] = EmptyStr then
     raise Exception.Create('엑셀로 변환할 자료가 존재하지 않습니다.');
end;
```
<br><br><br><br>





### 5. 엑셀 객체 생성 후 필드에 저장. <br>

- 엑셀 객체 생성
```delphi
function THB40010.Excel_module_class.make_excel() :Variant ;
var
 excel : Variant;
 i : integer;
 erro_message : PWideChar;
begin
  try
    { excel객체 생성 }
    excel := CreateOleObject('Excel.Application');

    { excel 새로운 페이지 생성 }
    excel.WorkBooks.Add;

    { excel sheet 추가 }
    for i := 0 to 1 Do //sheet 2개 추가 -> sheet 총 3개
    Begin
      excel.WorkSheets.add;
    End;

    Result := excel;

  except
    // 사용자 PC에 엑셀 프로그램이 설치되어 있지 않은 경우 처리.
    erro_message := 'Excel이 설치되어 있지 않습니다. 먼저 Excel을 설치하세요.';
    Application.MessageBox( erro_message ,'오류', MB_OK or MB_ICONERROR);
    Screen.Cursor  := crDefault;
    Exit;
  end;

end;
```
<br>

- 엑셀 객체 필드에 저장. 
```delphi
constructor THB40010.Excel_module_class.Create( temp_dept_name :String ; temp_grid1 , temp_grid2 :TAdvStringGrid );
var
   sheet_name , title :String;
begin
  생략...
  
  excel := make_excel();

  생략...
end;
```
<br><br><br><br>





### 6. 시작지점 설정. <br>

엑셀의 어느 셀부터 데이터를 넣을지 설정한다. <br>
```delphi
procedure THB40010.Excel_module_class.set_field_excel_start_end_cell();
var
 start_row , start_col , space_row , printPage_max_row :integer;
begin
  start_row :=3;
  start_col :=1;
  space_row :=1; //row 공백
  printPage_max_row := 24;

  생략...
  
end;
```
<br><br><br><br>





### 7. 시트이름 설정. <br>

- 시트명을 필드에 저장. <br>
```delphi
constructor THB40010.Excel_module_class.Create( temp_dept_name :String ; temp_grid1 , temp_grid2 :TAdvStringGrid );
var
   sheet_name , title :String;
begin
  생략...

  sheet_name := '휴가 계획서';
  
  생략...
end;  
```
<br>

- 엑셀객체를 사용해서 시트이름 설정 <br>
```delphi
procedure THB40010.Excel_module_class.set_excel_sheet_name( sheet_name:String );
begin
     excel.WorkSheets[1].name := sheet_name;
end;
```
<br><br><br><br>





### 8. 엑셀에 데이터 출력. <br>

엑셀객체의 range속성에 내용이 담긴 배열을 입력한다. <br>
엑셀객체와 배열은 필드에 저장된것을 사용한다. <br>

```delphi
procedure THB40010.Excel_module_class.insert_data_to_excel() ;
begin
     excel.Range[ excel_start_cell_1 , excel_end_cell_1  ].Value := grid1_data_arr;
     excel.Range[ excel_start_cell_2 , excel_end_cell_2  ].Value := grid2_data_arr;
end;
```
<br><br><br><br>





### 9. Cell 정렬 설정하기. <br>

모든 Cell을 동일하게 가운데 정렬 <br>

```delphi
procedure THB40010.Excel_module_class.set_aligment_center();
begin
     excel.Range[ excel_start_cell_1 , excel_end_cell_1 ].HorizontalAlignment := -4108;
     excel.Range[ excel_start_cell_2 , excel_end_cell_2 ].HorizontalAlignment := -4108;
end;
```
<br><br><br><br>





### 10. Cell 높이,넓이 설정하기. <br>

```delphi
procedure THB40010.Excel_module_class.set_height_width();
var
   I , printPage_max_row :integer;
   
begin
    printPage_max_row := 24;
      
    { 행 높이 조절 }
    for I := 2 to printPage_max_row * 2 do
    begin
         excel.WorkSheets[1].Rows[I].RowHeight := 20;
         excel.WorkSheets[1].Rows[I].Font.SIZE := 10; //글자크기
    end;
    { 열 넓이 조절 }
    for I := 2 to grid1_arr_col_length - 1  do
    begin
         excel.Cells.Item[ 1 , i ].ColumnWidth := 2.5;
    end;
end;
```
<br><br><br><br>





### 11. Cell 테두리 설정하기. <br>

```delphi
procedure THB40010.Excel_module_class.set_excel_line();
begin
  excel.Range[ excel_start_cell_1 , excel_end_cell_1 ].borders.lineStyle := 1;
  excel.Range[ excel_start_cell_2 , excel_end_cell_2 ].borders.lineStyle := 1;
end;
```
<br><br><br><br>





### 12. Cell 색 설정하기. <br>

주말 , 공휴일은 다른 색으로 구분한다. <br>

```delphi
procedure THB40010.Excel_module_class.set_color();
var
   col , printPage_max_row :integer;
begin
  printPage_max_row := 24;
  
  생략...

end;
```
<br><br><br><br>





### 13. Cell 글씨체 설정하기. <br>

```delphi
procedure THB40010.Excel_module_class.set_font();
begin
  excel.Range[ excel_start_cell_1 , excel_end_cell_1 ].font.Name := '나눔고딕';
  excel.Range[ excel_start_cell_2 , excel_end_cell_2 ].font.Name := '나눔고딕';
end;
```
<br><br><br><br>





### 14. 제목 설정하기. <br>


- 제목 필드에 저장. <br>
```delphi
constructor THB40010.Excel_module_class.Create( temp_dept_name :String ; temp_grid1 , temp_grid2 :TAdvStringGrid );
var
   sheet_name , title :String;
begin
  생략...

  title := dept_name + '  하기휴가 계획서';
  
  생략...
end; 
```
<br>

- 제목 설정. <br>
```delphi
procedure THB40010.Excel_module_class.set_cexcel_title( title:String );
var 
   printPage_max_row :integer;    
begin
   printPage_max_row := 24;
     
   excel.Cells[ 1 , 1 ] :=  title;
   set_title_design(1);
     
end;
```
<br><br><br><br>





### 마무리. <br>

- 단위별로 작업을 잘 나눴기 때문에 수정시에 필요한 코드만을 보고 해석 할 수 있게 되었다. <br>

- 자주 사용하는 객체나 변수를 필드로 두고 사용하니 중복 코드가 줄었다. <br> 

- naming 규칙을 체계적으로 정한다면 더 빨리 이해할 수 있는 코드가 될 것이다. <br>























<br><br><br><br>



```delphi

```
<br>















