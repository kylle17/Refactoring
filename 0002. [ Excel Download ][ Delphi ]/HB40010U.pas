unit HB40010U;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, FireDAC.Stan.Intf, DateUtils, Generics.Collections,
  FireDAC.Stan.Option, FireDAC.Stan.Param, FireDAC.Stan.Error, FireDAC.DatS,
  FireDAC.Phys.Intf, FireDAC.DApt.Intf, FireDAC.Stan.Async, FireDAC.DApt, QRCtrls, QuickRpt,
  Data.DB, FireDAC.Comp.DataSet, FireDAC.Comp.Client, Vcl.StdCtrls, Vcl.Buttons,
  Vcl.Imaging.jpeg, AdvUtil, Vcl.WinXPanels, Vcl.Grids, AdvObj, BaseGrid,
  AdvGrid, Vcl.ExtCtrls, Vcl.Menus, BASEU_HUR, Vcl.Imaging.pngimage, ComObj,
  AdvCustomComponent, AdvPDFIO, AdvGridPDFIO;

type
  TEinit = (clrAll, clrDetail, clrGrid);
  TEModeControl = (mcNone, mcAdd, mcMod);

  TWorkCode = class
    private
      WorkCode, CodeRef : string;
  end;

  TWorkType = class
    private
      BeginWorkTime, EndWorkTime : string;
  end;

  TMatchStr = class
    FMatchStr : string;
    constructor Create(matchStr:string);
  end;

type
  THB40010 = class(TBASE_HURF)
    Panel12: TPanel;
    Shape26: TShape;
    Shape28: TShape;
    Panel13: TPanel;
    lblTitleDetailGrid: TLabel;
    Panel8: TPanel;
    Shape13: TShape;
    Shape15: TShape;
    Panel9: TPanel;
    Image3: TImage;
    Label2: TLabel;
    ScrollBox2: TScrollBox;
    Shape1: TShape;
    Shape2: TShape;
    Shape3: TShape;
    Shape11: TShape;
    Splitter1: TSplitter;
    asgMaster: TAdvStringGrid;
    Shape27: TShape;
    Image2: TImage;
    Label3: TLabel;
    Label5: TLabel;
    cbxBFYear: TComboBox;
    pnlPrepareDetailGrid: TPanel;
    Label8: TLabel;
    cbxSDept: TComboBox;
    pnl1: TPanel;
    asgDetail: TAdvStringGrid;
    Panel5: TPanel;
    asgDetail2: TAdvStringGrid;
    popMenu: TPopupMenu;
    N00: TMenuItem;
    s1: TMenuItem;
    s2: TMenuItem;
    s3: TMenuItem;
    s4: TMenuItem;
    s5: TMenuItem;
    s6: TMenuItem;
    s7: TMenuItem;
    xxxxx: TMenuItem;
    p1: TMenuItem;
    p2: TMenuItem;
    p3: TMenuItem;
    p4: TMenuItem;
    p5: TMenuItem;
    p6: TMenuItem;
    p7: TMenuItem;
    p8: TMenuItem;
    p9: TMenuItem;
    p10: TMenuItem;
    p15: TMenuItem;
    p11: TMenuItem;
    p12: TMenuItem;
    p13: TMenuItem;
    p14: TMenuItem;
    N01: TMenuItem;
    N02: TMenuItem;
    N1: TMenuItem;
    N05: TMenuItem;
    N04: TMenuItem;
    N07: TMenuItem;
    N08: TMenuItem;
    N06: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    N15: TMenuItem;
    procedure SB_QUERYClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
//    procedure FormDestroy(Sender: TObject);
    procedure SB_INITClick(Sender: TObject);
    procedure SB_PRINTClick(Sender: TObject);
    procedure SB_EXCELClick(Sender: TObject);
    procedure cbxConDateChange(Sender: TObject);
    procedure ChangeEnrollInfo(Sender: TObject);
//    procedure WorkCodeClick(Sender: TObject);
//    procedure WorkTypeClick(Sender: TObject);
//    procedure sbCopyWPPreMonthClick(Sender: TObject);
    procedure sbDetailInitClick(Sender: TObject);
    procedure asgDetailGetAlignment(Sender: TObject; ARow, ACol: Integer;
      var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asgDetailGetCellColor(Sender: TObject; ARow, ACol: Integer;
      AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
    procedure asgDetailGetEditorType(Sender: TObject; ACol, ARow: Integer;
      var AEditor: TEditorType);
    procedure asgDetailKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
//    procedure asgDetailClipboardBeforePasteCell(Sender: TObject; ACol,
//     ARow: Integer; var Value: string; var Allow: Boolean);
//    procedure asgDetailCellValidate(Sender: TObject; ACol, ARow: Integer;
//      var Value: string; var Valid: Boolean);
//    procedure asgDetailClipboardCutDone(Sender: TObject; CellRect: TGridRect);
//    procedure sbAddNewWorkPlanClick(Sender: TObject);
//    procedure btnAppStsClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
//    procedure lblConfirmOrderClick(Sender: TObject);
    procedure asgDetailKeyPress(Sender: TObject; var Key: Char);
    procedure btnConfirmShortCutClick(Sender: TObject);
//    procedure cbxSDeptChange(Sender: TObject);
    procedure asgMasterClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure gP_GridToExcel_yj(const sgTemp, sgTemp2 : TStringGrid;
                         iColCount, iColCount2 : Integer;
                         const Title , Jogun : String ;
                         const asPos : array of integer;  //���ڿ�
                         const iFormat : integer = 1);


  private
    { Private declarations }
    FMode:TEModeControl;
    WorkCodeDict : TObjectDictionary<string, TWorkCode>;
    WorkTypeDict : TObjectDictionary<string, TWorkType>;
    WorkTypeShortCut:TDictionary<Char,Char>;
    HintShorCutList:TStringList;
    nowDateTime:TDateTime;
    FUserDeptCode     :string;
    FUserDeptOverCode :string;
    FUserDeptName     :string;
    FUserDeptOverName :string;
    FConfirmOrder     :string;

    { ��ü�����μ� 090: �������� }
    const
      FMngDept = '090';

    procedure Init(flag: TEinit; gridArr:array of TAdvStringGrid);
    procedure ModeControl(mode : TEModeControl; appSts:Boolean = True);
    procedure DetailQuery(Sender: TObject);

    procedure paint_weekend_color( excel_sheet:Variant ; i , iRowCnt , row :integer );
    procedure set_excel_cell_alignment_center( XL:Variant ; sheet_index , iStartRow , iRowCnt , iColCount: integer );


  public
    { Public declarations }

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

  end;// Excel_module_class ��



  end; //��

var
  HB40010: THB40010;

implementation

uses  MAINU, MainDMU, MASTER, MASTER_HUR, MASTER_CODE, HB40010P1U,
      POP_APPROVALU, POP_ORDERMSGU;

{$R *.dfm}

procedure THB40010.FormCreate(Sender: TObject);
var
  cWorkCode : TWorkCode;
  cWorkType : TWorkType;

  function CreateMenuItem(dispText:string; inputFlag:Integer = 0) : TMenuItem;
  var
    menuItem: TMenuItem;
  begin
    menuItem := TMenuItem.Create(popMenu);
    menuItem.AutoHotkeys := maParent;
    menuItem.Caption := dispText;
    Result := menuItem;
  end;
begin
inherited;

  SB_PRINT.Enabled := False;
  SB_SAVE.Enabled := False;
  SB_DELETE.Enabled := False;

  WorkCodeDict := TObjectDictionary<string, TWorkCode>.Create([doOwnsValues]);
  WorkTypeDict := TObjectDictionary<string, TWorkType>.Create([doOwnsValues]);
  { �ٹ����� ����Ű ��� }
  WorkTypeShortCut := TDictionary<Char,Char>.Create;
  { ����, �ٹ����� ����Ű ��Ʈ ��� }
  HintShorCutList := TStringList.Create;

  { �ٹ����� �Է� �˾��޴� ���� }
  popMenu.Items.Clear;
  popMenu.Items.Add(CreateMenuItem('����'));

  with Query1 do
  begin
    { �����ڵ� �߰� }
    Close;
    SQL.Clear;
    SQL.Add('SELECT CODECONT, CODEKEY2, CODEREF1   ');
    SQL.Add('  FROM HRBASECD                       ');
    SQL.Add(' WHERE CODEKEY1 = ''40''              ');
    SQL.Add('   AND CODEKEY2<> ''00''              ');
    SQL.Add('   AND DSPCHECK = ''Y''               ');
    SQL.Add(' ORDER BY DSPSEQNO                    ');
    Open;

    WorkCodeDict.Clear;
    while not Eof do
    begin
      cWorkCode := TWorkCode.Create;
      cWorkCode.WorkCode := Fields[1].AsString;
      cWorkCode.CodeRef  := Fields[2].AsString;

      WorkCodeDict.Add(Fields[0].AsString, cWorkCode);
      popMenu.Items.Add(CreateMenuItem(Fields[0].AsString, 40));
      Next;
    end;

    HintShorCutList.Add('�ٹ����� ����Ű ���'+#13);
    { �ٹ����� �߰� }
    Close;
    SQL.Clear;
    SQL.Add('SELECT CODEKEY2, CODEREF1 WORKBGIN, CODEREF2 WORKEND, CODEREF3 ');
    SQL.Add('  FROM HRBASECD                                         ');
    SQL.Add(' WHERE CODEKEY1 = ''42''                                ');
    SQL.Add('   AND CODEKEY2<> ''00''                                ');
    Open;

    WorkTypeDict.Clear;
    while not Eof do
    begin
      cWorkType := TWorkType.Create;
      { ����ٽð� }
      cWorkType.BeginWorkTime := Fields[1].AsString;
      cWorkType.EndWorkTime  := Fields[2].AsString;

      { ���� ����Ű ���� }
      if FieldByName('CODEREF3').AsString <> '' then
      begin
        WorkTypeShortCut.Add(Fields[3].AsString[1], Fields[0].AsString[1]);
        HintShorCutList.Add(Fields[0].AsString+ '   ' +Fields[1].AsString+'~'+Fields[2].AsString + ' : ' + Fields[3].AsString);
      end
      else
        HintShorCutList.Add(Fields[0].AsString+ '   ' +Fields[1].AsString+'~'+Fields[2].AsString + ' : ' + Fields[0].AsString);

      WorkTypeDict.Add(Fields[0].AsString, cWorkType);
      popMenu.Items.Items[0].Add(CreateMenuItem(Fields[0].AsString+ '   ' +Fields[1].AsString+'~'+Fields[2].AsString, 42));
      Next;
    end;

    { �α��� ���� �ٹ���ȹ�׷��ڵ� }
    Close;
    SQL.Clear;
    SQL.Add('SELECT B.WKPLNGRP, C.DEPTNAME, D.WKPLNGRP, E.DEPTNAME   ');
    SQL.Add('  FROM HREMPLMT A                                       ');
    SQL.Add('  LEFT JOIN HRDEPTCD B                                  ');
    SQL.Add('    ON B.DEPTCODE = A.DEPTCODE                          ');
    SQL.Add('   AND B.ENDDDATE = ''9999-99-99''                      ');
    SQL.Add('  LEFT JOIN HRDEPTCD C                                  ');
    SQL.Add('    ON C.DEPTCODE = B.WKPLNGRP                          ');
    SQL.Add('   AND C.ENDDDATE = ''9999-99-99''                      ');
    SQL.Add('  LEFT JOIN HRDEPTCD D                                  ');
    SQL.Add('    ON D.DEPTCODE = A.DEPTOVER                          ');
    SQL.Add('   AND D.ENDDDATE = ''9999-99-99''                      ');
    SQL.Add('  LEFT JOIN HRDEPTCD E                                  ');
    SQL.Add('    ON E.DEPTCODE = D.WKPLNGRP                          ');
    SQL.Add('   AND E.ENDDDATE = ''9999-99-99''                      ');
    SQL.Add(' WHERE A.EMPLNUMB = :EMPLNUMB                           ');
    ParamByName('EMPLNUMB').AsString := MASTER_USERID;
    Open;

    { ���� �ٹ���ȹǥ �μ��׷��ڵ�(���μ�, ���Ӻμ�) }
    FUserDeptCode     := Fields[0].AsString;
    FUserDeptName     := Fields[1].AsString;
    FUserDeptOverCode := Fields[2].AsString;
    FUserDeptOverName := Fields[3].AsString;
  end;

  { default alclient }
  pnlPrepareDetailGrid.Align := alClient;

  { ���α׷� ����� �� ��¥ ���� }
  nowDateTime := StrToDateTime(SERVER_DATETIME);
  if DayOf(nowDateTime) > 31 then
    nowDateTime := IncYear(nowDateTime);

  { �޺��ڽ� ������ ���� }
  SetCbxAddYearList(cbxBFYear, 0, 3);

  { �׸��庰 ���� ���� }
  SetGridSortSetting(asgMaster);

  { �����̺�Ʈ ���� }
  CommonEventBinding;
end;
procedure THB40010.FormShow(Sender: TObject);
begin
  inherited;
  { ��ü �ʱ�ȭ }
  SB_INITClick(nil);
end;

procedure THB40010.FormActivate(Sender: TObject);
begin
  inherited;
  { �� Ȱ��ȭ �� }
  if asgMaster.Cells[0, asgMaster.FixedRows] <> '' then
    SB_QUERYClick(nil);

  if (asgDetail.Cells[0, asgDetail.FixedRows] <> '') and (FMode = mcMod) then
    DetailQuery(nil);

  if (asgDetail2.Cells[0, asgDetail2.FixedRows] <> '') and (FMode = mcMod) then
    DetailQuery(nil);

end;

{procedure THB40010.FormDestroy(Sender: TObject);
var
  i:Integer;
begin
  inherited;
  for i := 0 to cbxDept.Items.Count - 1 do
    cbxDept.Items.Objects[i].Free;

  WorkCodeDict.Free;
  WorkTypeDict.Free;
  WorkTypeShortCut.Free;
  HintShorCutList.Free;
end;
 }
procedure THB40010.SB_INITClick(Sender: TObject);
begin
  inherited;
  Init(clrAll, [asgMaster, asgDetail]);
  Init(clrAll, [asgMaster, asgDetail2]);
end;

procedure THB40010.SB_QUERYClick(Sender: TObject);
var
  rowSeq, i: Integer;
  bfMonth:string;
begin
  try
    SB_QUERY.Enabled := False;
    asgMaster.BeginUpdate;
    try
      if (not IsDate(cbxBFYear.Text+'-'+'07'+'-'+'01')) then
        raise Exception.Create('[��ȸ����] ��ȹ����� ���������� �Է���� Ȯ�ιٶ��ϴ�.');

      { ��ȸ���� }
      //bfMonth := EncodeDate( StrToInt(cbxBFYear.Text));

      { ��ȸ �� �ʱ�ȭ }
      Init(clrGrid, [asgMaster]);

      with Query1, asgMaster do
      begin
        Close;
        SQL.Clear;
        SQL.Add('  SELECT DEPTNAME, DEPTCODE, STATDATE, ENDDDATE                                                   ');
        SQL.Add('    FROM HRDEPTCD                                                                                 ');
        SQL.Add('   WHERE SUBSTR(ENDDDATE, 1, 7) > :BFMONTH                                                        ');
        SQL.Add('     AND DEPTCODE IN (''090'', ''610'', ''310'', ''222'', ''210'', ''240'', ''221'', ''400'')     ');
        //if(cbxSDept.ItemIndex <> 0) then
        //SQL.Add('     AND A.DEPTCODE = :DEPTCODE                                                                   ');

        { Setting Parameter }
        ParamByName('BFMONTH').AsString := cbxBFYear.Text+'-07';
        //ShowMessage(cbxBFYear.Text+'-07');
        //if(cbxSDept.ItemIndex <> 0) then
        //ParamByName('DEPTCODE').AsString := TMatchStr(cbxSDept.Items.Objects[cbxSDept.ItemIndex]).FMatchStr;
        Open;
        if Query1.IsEmpty then
          Exit;

        rowSeq := FixedRows;

        while not Eof do
        begin
          Ints[0, rowSeq] := rowSeq;
          Cells[1, rowSeq] := FieldByName('DEPTNAME').AsString;

          Inc(rowSeq);
          Next;
        end;
        RowCount := rowSeq;

      end;
    except
      on E: Exception do
      begin
        CALL_ERROR('��ȸ ����', Label2.Caption + ' ��ȸ �� ������ �߻��߽��ϴ�.' + #13#10#13#10 + E.Message);
        Exit;
      end;
    end;
  finally
    asgMaster.EndUpdate;
    sb_query.Enabled := True;
  end;
end;

procedure THB40010.SB_PRINTClick(Sender: TObject);
var
  printFormClass:TFormClass;
  printForm : TForm;
  quickRep: TQuickRep;
begin
  inherited;
  try
    //SB_PRINT.Enabled := False;
    try
      { ��ȿ���˻� }
      if cbxSDept.ItemIndex < 0 then
        raise Exception.Create('7,8�� �ϱ��ް���ȹ�� ����� ���� �� �������ּ���.');

      printFormClass := THB40010P1;

      printForm := printFormClass.Create(Self);
      try
        quickRep := TQuickRep(printForm.FindComponent('QuickRep'));
        quickRep.PreviewInitialState := wsMaximized;
        quickRep.PreviewModal;
      finally
        printForm.Free;
      end;
    except
      on E: Exception do
      begin
        CALL_ERROR('��¿���', '�ڷ������ ���ܻ����� �߻��߽��ϴ�. ' + #13#10#13#10 + E.Message);
      end;
    end;
  finally
    SB_PRINT.Enabled := True;
  end;
end;


procedure THB40010.SB_EXCELClick(Sender: TObject);
var
  grid1 , grid2 :TAdvStringGrid;
  convertToStrArr: array of Integer;
  excel_module :Excel_module_class;
  dept_name :String;
begin
  inherited;
  try
    SB_EXCEL.Enabled := False;
    try

      grid1 := asgDetail;
      grid2 := asgDetail2;

      dept_name := cbxSDept.Text;

      excel_module := Excel_module_class.create( dept_name , grid1 , grid2 );
      excel_module.free;


      { ���� ���� }
//      SetLength(convertToStrArr, grid1.ColCount - 1);
//      gP_GridToExcel_yj(grid1, grid2, grid1.ColCount, grid2.ColCount, fmTITLE.Caption, cbxSDept.Text, convertToStrArr);

    except
      on E: Exception do
      begin
        CALL_ERROR('������ȯ����', '������ȯ�� ���ܻ����� �߻��߽��ϴ�. ' + #13#10#13#10 + E.Message);
      end;
    end;
  finally
    SB_EXCEL.Enabled := True;
  end;
end;

{===================================================================================================================}
{=====                        ======================================================================================}
{=====      Excel_Moduel      ======================================================================================}
{=====                        ======================================================================================}
{===================================================================================================================}
constructor THB40010.Excel_module_class.Create( temp_dept_name :String ; temp_grid1 , temp_grid2 :TAdvStringGrid );
var
   sheet_name , title :String;
begin
  inherited Create;

  dept_name := temp_dept_name;

  grid1 := temp_grid1;
  grid2 := temp_grid2;

  grid_is_empty();

  excel := make_excel();

  set_excel_print_page();

  set_field_grid_data_arr();
  set_field_grid1_arr_length();
  set_field_excel_start_end_cell();

  sheet_name := '�ϱ��ް� ��ȹ��';
  set_excel_sheet_name( sheet_name );

  insert_data_to_excel();
  set_aligment_center();
  set_height_width();
  set_excel_line();
  set_color();
  set_font();

  title := dept_name + '  �ϱ��ް� ��ȹ��';
  set_cexcel_title( title);

  excel.Visible:= True;

end;

procedure THB40010.Excel_module_class.grid_is_empty();
begin
  if grid1.Cells[0, grid1.FixedRows] = EmptyStr then
     raise Exception.Create('������ ��ȯ�� �ڷᰡ �������� �ʽ��ϴ�.');
end;

function THB40010.Excel_module_class.make_excel() :Variant ;
var
 excel : Variant;
 i : integer;
 erro_message : PWideChar;
begin
  try
    { excel��ü ���� }
    excel := CreateOleObject('Excel.Application');

    { excel ���ο� ������ ���� }
    excel.WorkBooks.Add;

    { excel sheet �߰� }
    for i := 0 to 1 Do //sheet 2�� �߰� -> sheet �� 3��
    Begin
      excel.WorkSheets.add;
    End;

    Result := excel;

  except
    erro_message := 'Excel�� ��ġ�Ǿ� ���� �ʽ��ϴ�. ���� Excel�� ��ġ�ϼ���.';
    Application.MessageBox( erro_message ,'����', MB_OK or MB_ICONERROR);
    Screen.Cursor  := crDefault;
    Exit;
  end;

end;

procedure THB40010.Excel_module_class.set_excel_print_page();
begin
     excel.WorkSheets[1].PageSetup.Orientation := $00000002; //�������� ����(1) , ����(2)
     excel.WorkSheets[1].PageSetup.Zoom := False;
     excel.WorkSheets[1].PageSetup.FitToPagesWide := 1;
     excel.WorkSheets[1].PageSetup.FitToPagesTall := False;
end;

procedure THB40010.Excel_module_class.set_field_grid_data_arr();
begin
     grid1_data_arr := make_grid_data_arr( grid1 );
     grid2_data_arr := make_grid_data_arr( grid2 );
end;

procedure THB40010.Excel_module_class.set_field_grid1_arr_length();
begin
  grid1_arr_row_length := get_grid_arr_row_length( grid1_data_arr );
  grid1_arr_col_length := get_grid_arr_col_length( grid1_data_arr );
  grid2_arr_row_length := get_grid_arr_row_length( grid2_data_arr );
  grid2_arr_col_length := get_grid_arr_col_length( grid2_data_arr );
end;

procedure THB40010.Excel_module_class.set_field_excel_start_end_cell();
var
 start_row , start_col , space_row , printPage_max_row :integer;
begin
  start_row :=3;
  start_col :=1;
  space_row :=1; //row ����
  printPage_max_row := 24;

  if dept_name <> '��������' then
  begin
    excel_start_cell_1 := excel.WorkSheets[1].Cells[ start_row                                                               ,   start_col               ];
    excel_end_cell_1   := excel.WorkSheets[1].Cells[ start_row - space_row + grid1_arr_row_length                            ,   grid1_arr_col_length    ];
          
    excel_start_cell_2 := excel.WorkSheets[1].Cells[ start_row + space_row + grid1_arr_row_length                            ,   start_col               ];
    excel_end_cell_2   := excel.WorkSheets[1].Cells[ start_row + grid1_arr_row_length + grid2_arr_row_length                 ,   grid2_arr_col_length    ];
  end;

  if dept_name = '��������' then
  begin
    excel_start_cell_1 := excel.WorkSheets[1].Cells[ start_row                                                               ,   start_col               ];
    excel_end_cell_1   := excel.WorkSheets[1].Cells[ start_row - space_row + grid1_arr_row_length                            ,   grid1_arr_col_length    ];

    excel_start_cell_2 := excel.WorkSheets[1].Cells[ printPage_max_row + start_row - space_row                               ,   start_col               ];
    excel_end_cell_2   := excel.WorkSheets[1].Cells[ printPage_max_row + start_row - (space_row*2) + grid2_arr_row_length    ,   grid2_arr_col_length    ];  
  end;
  
end;

procedure THB40010.Excel_module_class.set_excel_sheet_name( sheet_name:String );
begin
     excel.WorkSheets[1].name := sheet_name;
end;

procedure THB40010.Excel_module_class.set_cexcel_title( title:String );
var 
   printPage_max_row :integer;    
begin
   printPage_max_row := 24;
     
   {���� ����}
   excel.Cells[ 1 , 1 ] :=  title;
   set_title_design(1);
     
   {�������� �����߰�}
   if dept_name = '��������' then 
   begin
     excel.Cells[ printPage_max_row , 1 ] :=  title;
     set_title_design( printPage_max_row );
   end;

end;

procedure THB40010.Excel_module_class.set_title_design( row :integer );
var
   start_cell , end_cell :Variant;
begin
   start_cell := excel.WorkSheets[1].Cells[ row , 1 ];
   end_cell   := excel.WorkSheets[1].Cells[ row , grid1_arr_col_length ];
   {�� ����}
   excel.Range[ start_cell , end_cell ].Merge;
   {�� ����}
   excel.WorkSheets[1].Rows[row].RowHeight := 50;
   {��� ����}
   excel.Range[ start_cell , end_cell ].HorizontalAlignment := -4108;
   {����ũ��}
   excel.Range[ start_cell , end_cell ].Font.SIZE := 25;
   {�۾�ü}
   excel.Range[ start_cell , end_cell ].font.Name := '�������';
   {����}
   excel.Range[ start_cell , end_cell ].font.bold := True;
end;

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

function THB40010.Excel_module_class.get_grid_arr_row_length( grid_data_arr :OleVariant ) :integer;
begin
     result := VarArrayHighBound( grid_data_arr , 1 );
end;

function THB40010.Excel_module_class.get_grid_arr_col_length( grid_data_arr :OleVariant ) :integer;
begin
     result := VarArrayHighBound( grid_data_arr , 2 );
end;

procedure THB40010.Excel_module_class.insert_data_to_excel() ;
begin
     excel.Range[ excel_start_cell_1 , excel_end_cell_1  ].Value := grid1_data_arr;
     excel.Range[ excel_start_cell_2 , excel_end_cell_2  ].Value := grid2_data_arr;
end;

procedure THB40010.Excel_module_class.set_aligment_center();
begin
     excel.Range[ excel_start_cell_1 , excel_end_cell_1 ].HorizontalAlignment := -4108;
     excel.Range[ excel_start_cell_2 , excel_end_cell_2 ].HorizontalAlignment := -4108;
end;

procedure THB40010.Excel_module_class.set_height_width();
var
   I , printPage_max_row :integer;
   
begin
    printPage_max_row := 24;
      
    { �� ���� ���� }
    for I := 2 to printPage_max_row * 2 do
    begin
         excel.WorkSheets[1].Rows[I].RowHeight := 20;
         excel.WorkSheets[1].Rows[I].Font.SIZE := 10; //����ũ��
    end;
    { �� ���� ���� }
    for I := 2 to grid1_arr_col_length - 1  do
    begin
         excel.Cells.Item[ 1 , i ].ColumnWidth := 2.5;
    end;
end;

procedure THB40010.Excel_module_class.set_excel_line();
begin
  excel.Range[ excel_start_cell_1 , excel_end_cell_1 ].borders.lineStyle := 1;
  excel.Range[ excel_start_cell_2 , excel_end_cell_2 ].borders.lineStyle := 1;
end;

procedure THB40010.Excel_module_class.set_color();
var
   col , printPage_max_row :integer;
begin
  printPage_max_row := 24;
       
  { ��,�� ���� ���� }
  For col := 1 to grid1_arr_col_length do
  begin
      { grid1 }
      // �� , ��      
      if (excel.WorkSheets[1].Cells[ 4 , col ].Value = '��') OR (excel.WorkSheets[1].Cells[ 4 , col ].Value = '��') then
      begin
        excel.WorkSheets[1].Cells[ 4 , col ].Interior.Color := $DDDBF8;
        excel.WorkSheets[1].Cells[ 3 , col ].Interior.Color := $DDDBF8;
      end;      

      
      { grid2 }
      // �� , ��
      if (excel.WorkSheets[1].Cells[ grid1_arr_row_length + 5 , col ].Value = '��') OR (excel.WorkSheets[1].Cells[ grid1_arr_row_length + 5 , col ].Value = '��') then
      begin
        excel.WorkSheets[1].Cells[ grid1_arr_row_length+5 , col ].Interior.Color := $DDDBF8;
        excel.WorkSheets[1].Cells[ grid1_arr_row_length+4 , col ].Interior.Color := $DDDBF8;
      end;
      // ������ , �ΰ� â�������
      if (excel.WorkSheets[1].Cells[ grid1_arr_row_length + 4 , col ].Value = '15') OR (excel.WorkSheets[1].Cells[ grid1_arr_row_length + 4 , col ].Value = '30') then
      begin
        excel.WorkSheets[1].Cells[ grid1_arr_row_length+5 , col ].Interior.Color := $00CEEFFE;
        excel.WorkSheets[1].Cells[ grid1_arr_row_length+4 , col ].Interior.Color := $00CEEFFE;
      end;

      
      { grid2 (���������� ���) }
      // �� , ��
      if (excel.WorkSheets[1].Cells[ printPage_max_row + 3 , col ].Value = '��') OR (excel.WorkSheets[1].Cells[ printPage_max_row + 3 , col ].Value = '��') then
      begin
        excel.WorkSheets[1].Cells[ printPage_max_row + 3 , col ].Interior.Color := $DDDBF8;
        excel.WorkSheets[1].Cells[ printPage_max_row + 2 , col ].Interior.Color := $DDDBF8;
      end;
      // ������ , �ΰ� â�������
      if (excel.WorkSheets[1].Cells[ printPage_max_row + 2 , col ].Value = '15') OR (excel.WorkSheets[1].Cells[ printPage_max_row + 2 , col ].Value = '30') then
      begin
        excel.WorkSheets[1].Cells[ printPage_max_row + 3 , col ].Interior.Color := $00CEEFFE;
        excel.WorkSheets[1].Cells[ printPage_max_row + 2 , col ].Interior.Color := $00CEEFFE;
      end;
  end;
end;

procedure THB40010.Excel_module_class.set_font();
begin
  excel.Range[ excel_start_cell_1 , excel_end_cell_1 ].font.Name := '�������';
  excel.Range[ excel_start_cell_2 , excel_end_cell_2 ].font.Name := '�������';
end;




{------------------------------------------------------------------------------
   �׸��� ������ȯ-1
 ------------------------------------------------------------------------------}
procedure THB40010.gP_GridToExcel_yj(const sgTemp, sgTemp2 : TStringGrid;
                         iColCount, iColCount2 : Integer;
                         const Title , Jogun : String ;
                         const asPos : array of integer;  //���ڿ�
                         const iFormat : integer = 1);
CONST
  excel_max_row = 1048572 ; //�������� �����ϴ� �ִ� ���.
var
  iSheetCnt : integer; //��Ʈ��.
  iStartRow : integer; //������ ���۵Ǵ� ���� �����Ѵ�.
  XL, ArrT,ArrT2: Variant;
  i,j,z, a,v : integer;
  ArrV, ArrV2 ,Format: OleVariant;
  Row,Col : Integer;
  iRowCnt : integer;
  sTemp, sTemp2 : String;
begin
  iStartRow := 4;
  Screen.Cursor:= crHourGlass;

  try
    XL := CreateOleObject('Excel.Application');
  except
    Application.MessageBox('Excel�� ��ġ�Ǿ� ���� �ʽ��ϴ�. ���� Excel�� ��ġ�ϼ���.',
      '����', MB_OK or MB_ICONERROR);
    Screen     .Cursor  := crDefault;
    Exit;
  end;

  CALL_PROGRESS(1);

  //���ο� ������ ����
  XL.WorkBooks.Add;

  //�� ��Ʈ ���� ����Ѵ�.
  iSheetCnt := ((sgTemp.RowCount + sgTemp2.RowCount)  div excel_max_row) +1; //
  //��Ʈ�� 3���� ������ ��Ʈ�� �߰��Ѵ�.
  for i := 2 to iSheetCnt Do
  Begin
    XL.WorkSheets.add;
  End;


  icolCount := asgDetail.ColCount;
  iColCount2 :=  asgDetail2.ColCount;



  Try
    For i := 1 to iSheetCnt do
    Begin
         //Rowcount ��� �� SHEET�̸� ����.
         if i = iSheetCnt Then
          Begin
                iRowCnt := sgTemp.RowCount-((i-1)*excel_max_row);
          End
          Else
         Begin
               iRowCnt := excel_max_row;
         End;

          //��Ʈ�̸��� ����� ����.
         XL.WorkSheets[i].name := '�ϱ��ް���ȹ��'+IntToStr(i);

          //��Ʈ�� 2�� �̻��̸� �������� ���� �Ҵ��Ѵ�.
          If iSheetCnt > 1 Then
          Begin
            ArrT := VarArrayCreate([0, iColCount-1], VarVariant);
            ArrT2 := VarArrayCreate([0, iColCount2-1], VarVariant);
          for  Col := 0 to  iColCount -1 do
          Begin
            // ArrT[Col] := sgTemp.Cells[Col,0];
            ArrT2[Col] := sgTemp2.Cells[Col, 0];
          End;
         End;

          //������°�.
    	  ArrV := VarArrayCreate([0, iRowCnt-1, 0, iColCount -1 ], VarVariant);
          For Row := 0 to iRowCnt -1 do
          begin
            For Col := 0 to iColCount - 1 do
                begin
                  Begin
                         sTemp := '';
                         if (col = asPos[j]) Then //���ڿ��̸�.
                             Begin
                              sTemp := '''';
                              Break;
                            end ;
                  End;

                  ArrV[Row, Col] := sTemp + sgTemp.cells[Col, Row+ ((i-1) * excel_max_row)];
                end;
          end;

          ArrV2 := VarArrayCreate([0, iRowCnt-1, 0, iColCount2 -1 ], VarVariant);
          For Row := 0 to iRowCnt -1 do
          begin
            For Col := 0 to iColCount2 - 1 do
                begin
                  For j:= 0 to Length(asPos) -1 do
                  Begin
                         sTemp2 := '';
                         if (col = asPos[j]) Then //���ڿ��̸�.
                             Begin
                              sTemp2 := '''';
                              Break;
                             end ;
                  End;

                  ArrV2[Row, Col] := sTemp2 + sgTemp2.cells[Col, Row+ ((i-1) * excel_max_row)];
                end;
          end;


          //����� ��ȸ ������ �ִ´�.
          //'���' ����
          XL.Range[XL.WorkSheets[i].Cells[iStartRow, 33], XL.WorkSheets[i].Cells[iStartRow+1, 33]].Merge;
          if (iRowCnt+iStartRow)*2-4 <25 then
            XL.Range[XL.WorkSheets[i].Cells[iStartRow+iRowCnt+1, 33], XL.WorkSheets[i].Cells[iStartRow+iRowCnt+2, 33]].Merge
          else
            XL.Range[XL.WorkSheets[i].Cells[29, 33], XL.WorkSheets[i].Cells[30, 33]].Merge;


           //����Ÿ �ֱ�
           XL.Cells[1,1] :=  Jogun +'  '+ Title;

           //����
           XL.Range[XL.WorkSheets[i].Cells[1, 1], XL.WorkSheets[i].Cells[1, iColCount]].Merge;
           XL.Range[XL.WorkSheets[i].Cells[2, 1], XL.WorkSheets[i].Cells[2, iColCount]].Merge;

           XL.WorkSheets[i].Rows[1].RowHeight := 50;
           XL.WorkSheets[i].Rows[1].Font.SIZE := 25;
           XL.WorkSheets[i].Rows[1].HorizontalAlignment := -4108;


          if (iRowCnt+iStartRow)*2-4 <25 then
          begin
            for z := iStartRow to (iRowCnt+iStartRow)*2-4 do
            begin
              XL.WorkSheets[i].Rows[z].RowHeight := 25;
              XL.WorkSheets[i].Rows[z].Font.SIZE := 15;
            end;

          end
          else
          begin
            //����Ÿ �ֱ�
            XL.Cells[26,1] :=  Jogun +'  '+ Title;

            //����
            XL.Range[XL.WorkSheets[i].Cells[26, 1], XL.WorkSheets[i].Cells[26, iColCount]].Merge;
            XL.Range[XL.WorkSheets[i].Cells[27, 1], XL.WorkSheets[i].Cells[27, iColCount]].Merge;

            for z := iStartRow-2 to (iRowCnt+iStartRow)*2+5 do
            begin
              XL.WorkSheets[i].Rows[z].RowHeight := 25;
              XL.WorkSheets[i].Rows[z].Font.SIZE := 15;
            end;

           XL.WorkSheets[i].Rows[26].RowHeight := 50;
           XL.WorkSheets[i].Rows[26].Font.SIZE := 30;
           XL.WorkSheets[i].Rows[26].HorizontalAlignment := -4108;

          end;

          If  i > 1 Then //ù��° ��Ʈ�� �ƴϸ� ������ ���� �߰��Ѵ�.
            Begin
              XL.Range[XL.WorkSheets[i].Cells[iStartRow, 1], XL.WorkSheets[i].Cells[iStartRow, iColCount]].Value := ArrT2;
              XL.Range[XL.WorkSheets[i].Cells[iStartRow+1, 1], XL.WorkSheets[i].Cells[iRowCnt+iStartRow-1, iColCount]].Value := ArrV;
            end
          Else
          Begin

            if (iRowCnt+iStartRow)*2-4 <25 then
            begin
                    XL.Range[XL.WorkSheets[i].Cells[iStartRow, 1], XL.WorkSheets[i].Cells[iRowCnt+iStartRow-1, iColCount]].Value := ArrV;
              XL.Range[XL.WorkSheets[i].Cells[iStartRow+iRowCnt+1, 1], XL.WorkSheets[i].Cells[(iRowCnt+iStartRow)*2-4, iColCount2]].Value := ArrV2;
            end
            else
            begin
              XL.Range[XL.WorkSheets[i].Cells[iStartRow, 1], XL.WorkSheets[i].Cells[iRowCnt+iStartRow-1, iColCount]].Value := ArrV;
              XL.Range[XL.WorkSheets[i].Cells[29, 1], XL.WorkSheets[i].Cells[(iRowCnt+iStartRow)*2+5, iColCount]].Value := ArrV2
            end;
          End;



          //���� ����.
          if (iRowCnt+iStartRow)*2-4 <25 then
          begin
            XL.Range[XL.WorkSheets[i].Cells[iStartRow+iRowCnt-1, 1], XL.WorkSheets[i].Cells[iStartRow, iColCount]].borders.lineStyle := 1;
            XL.Range[XL.WorkSheets[i].Cells[iStartRow+iRowCnt+1, 1], XL.WorkSheets[i].Cells[(iRowCnt+iStartRow)*2-4, iColCount]].borders.lineStyle := 1;
          end
          else
          begin
            XL.Range[XL.WorkSheets[i].Cells[iStartRow+iRowCnt-1, 1], XL.WorkSheets[i].Cells[iStartRow, iColCount]].borders.lineStyle := 1;
            XL.Range[XL.WorkSheets[i].Cells[29, 1], XL.WorkSheets[i].Cells[(iRowCnt+iStartRow)*2+5, iColCount]].borders.lineStyle := 1;
          end;


         // ��� ����
         set_excel_cell_alignment_center( XL , i , iStartRow , iRowCnt , iColCount );


         // ��,�� cell ������ ĥ�ϱ�.
         paint_weekend_color( XL.WorkSheets[i] , i , iRowCnt , iStartRow );



         //���ڿ� ���� ����.
         if iFormat = 1 then
            Format := '#,##0;[����]-#,##0;_-* "-"???_-;_-@_-'
         else
         if iFormat = 2 then
            Format := '#,##0.#0;[����]-#,##0.#0;_-* "-"???_-;_-@_-';

	 // XL.Range[XL.WorkSheets[i].Cells[iStartRow, 1], XL.WorkSheets[i].Cells[iRowCnt+iStartRow-1, iColCount]].NumberFormatLocal := Format;
         // XL.Range[XL.WorkSheets[i].Cells[iStartRow+iRowCnt+1, 1], XL.WorkSheets[i].Cells[(iRowCnt+iStartRow)*2-4, iColCount]].NumberFormatLocal := Format;

         XL.WorkSheets[i].Columns.Autofit;
    End;



    CALL_PROGRESS(9);
    XL.Visible:= True;
    // XL.WorkbookS[1].SaveAs('C:\TEST.XLS');
    // XL.WorkbookS[1].Password := '1001'; //��ȣ����.
  finally
     Screen.Cursor  := crDefault;
     XL := unAssigned;
  end;
End;


procedure THB40010.paint_weekend_color( excel_sheet:Variant ; i , iRowCnt , row :integer );
var
   Col :integer;
begin
   for Col := 1 to asgDetail.ColCount do
   begin
      if (excel_sheet.Cells[row+1, Col].Value = '��') OR  (excel_sheet.Cells[row+1, Col].Value = '��') then
      begin
          excel_sheet.Cells[row, Col].Interior.Color := $DDDBF8;
          excel_sheet.Cells[row+1, Col].Interior.Color := $DDDBF8;
      end;


      if (iRowCnt+row)*2-4 <25 then
      begin

       if (excel_sheet.Cells[iRowCnt+row+2, Col].Value = '��') OR (excel_sheet.Cells[iRowCnt+row+2, Col].Value = '��') then
        begin
          excel_sheet.Cells[iRowCnt+row+2, Col].Interior.Color := $DDDBF8;
          excel_sheet.Cells[iRowCnt+row+1, Col].Interior.Color := $DDDBF8;
        end;
      end
      else
      begin
       if (excel_sheet.Cells[30, Col].Value = '��') OR (excel_sheet.Cells[30, Col].Value = '��') then
        begin
          excel_sheet.Cells[30, Col].Interior.Color := $DDDBF8;
          excel_sheet.Cells[29, Col].Interior.Color := $DDDBF8;
        end;
      end;
   end;
end;


procedure THB40010.set_excel_cell_alignment_center( XL:Variant ; sheet_index , iStartRow , iRowCnt , iColCount: integer );
var
   excel_range:Variant;
   start_cell , end_cell : Variant;
begin
  if (iRowCnt+iStartRow)*2-4 <25 then
  begin
    start_cell  := XL.WorkSheets[sheet_index].Cells[iStartRow+iRowCnt-1,1];
    end_cell :=  XL.WorkSheets[sheet_index].Cells[iStartRow, iColCount];
    XL.Range[ start_cell , end_cell ].HorizontalAlignment := -4108;


    start_cell  := XL.WorkSheets[sheet_index].Cells[iStartRow+iRowCnt+1,1];
    end_cell :=  XL.WorkSheets[sheet_index].Cells[(iRowCnt+iStartRow)*2-4, iColCount];
    XL.Range[ start_cell , end_cell  ].HorizontalAlignment := -4108;
  end
  else
  begin
    start_cell := XL.WorkSheets[sheet_index].Cells[iStartRow+iRowCnt-1,1];
    end_cell := XL.WorkSheets[sheet_index].Cells[iStartRow, iColCount];
    XL.Range[ start_cell , end_cell ].HorizontalAlignment := -4108;


    start_cell := XL.WorkSheets[sheet_index].Cells[iStartRow+iRowCnt+1,1];
    end_cell := XL.WorkSheets[sheet_index].Cells[(iRowCnt+iStartRow)*2+5, iColCount];
    XL.Range[ start_cell , end_cell].HorizontalAlignment := -4108;
  end;
end;

procedure THB40010.Init(flag: TEinit; gridArr: array of TAdvStringGrid);
var
  i:Integer;
  sDate:TDateTime;
begin
  if Length(gridArr) > 0 then
  begin
    for i := Low(gridArr) to High(gridArr) do
    begin
      { Ǫ�� ���� }
      gridArr[i].FloatingFooter.Visible := False;
      gridArr[i].ClearSelectedCells;
      gridArr[i].OwnsObjects := True;
      if gridArr[i] = asgDetail then
      begin
        ASG_CLEAR(gridArr[i], 0, 0, 0, 0);
//        lblAppStsName.Font.Color := $00843123;
//        lblAppStsMsg.Visible := False;
//        lblConfirmOrder.Visible := False;
        FConfirmOrder := EmptyStr;
      end
      else if gridArr[i] = asgDetail2 then
      begin
        ASG_CLEAR(gridArr[i], 0, 0, 0, 0);
        FConfirmOrder := EmptyStr;
      end
      else
        ASG_CLEAR(gridArr[i], 0, gridArr[i].FixedRows, 0, 0);

      { ���ĵ�������� ���Ʈ }
      if gridArr[i].SortIndexes.Count > 0 then
        gridArr[i].QUnSort;

      { �ʱ�ȭ�� selection hide }
      gridArr[i].AutoHideSelection := True;
    end;
  end;
  if flag in [clrAll, clrDetail] then
  begin
    cbxBFYear.ItemIndex := cbxBFYear.Items.IndexOf(IntToStr(YearOf(nowDateTime)));
    cbxSDept.ItemIndex := -1;

    ModeControl(mcNone);
  end;

  if flag in [clrAll] then
  begin
//    if cbxSDept.Enabled then
//      cbxSDept.ItemIndex := 0;

    {�⺻����}
    sDate := IncYear(sDate, -2);
    cbxBFYear.ItemIndex := cbxBFYear.Items.IndexOf(IntToStr(YearOf(nowDateTime)));

    { ������� �����̺�Ʈ }
    ChangeEnrollInfo(nil);

    { ��¥ �����̺�Ʈ }
    cbxConDateChange(nil);

    cbxBFYear.SetFocus;
  end;
end;

procedure THB40010.asgMasterClickCell(Sender: TObject; ARow, ACol: Integer);
var
  selMonth:TDateTime;
  selDept:string;
begin
  if (ARow = 0) or (asgMaster.Cells[1, ARow] = EmptyStr) then Exit;

  { selection control }
  if asgMaster.AutoHideSelection then
    asgMaster.AutoHideSelection := False;

    selDept := asgMaster.Cells[1, ARow];

  { ������� ���� }
  cbxSDept.ItemIndex := cbxSDept.Items.IndexOf(selDept);
  ChangeEnrollInfo(Sender);
end;


procedure THB40010.DetailQuery(Sender: TObject);
var
  rowSeq, colSeq, i, c: Integer;
  plnMonth, bfDate, afDate, workDate: TDateTime;
  deptCode, wkCodeStr:string;
  WorkType:TWorkType;
  WorkCode:TWorkCode;
  canModWkPln:Boolean;
  BFMONTH, AFMONTH : string;
begin
  try
    asgDetail.BeginUpdate;
    asgDetail2.BeginUpdate;
    pnlPrepareDetailGrid.Visible := False;
    asgDetail.Visible := true;
    asgDetail2.Visible := true;
    pnl1.Visible := true;
    panel5.Visible := true;

    try
     if not IsDate(cbxBFYear.Text + '-' + '07' + '-' + '01') then
        raise Exception.Create('[��ȹ���] �׸��� ���������� �Է���� Ȯ�ιٶ��ϴ�.');


      plnMonth := EncodeDate(StrToInt(cbxBFYear.Text), 7, 1);
      bfDate := EncodeDate(YearOf(IncMonth(plnMonth, -1)), MonthOf(IncMonth(plnMonth, -1)), 21);
      deptCode := TMatchStr(cbxSDept.Items.Objects[cbxSDept.ItemIndex]).FMatchStr;
      { �ʱ�ȭ }
      Init(clrGrid, [asgDetail]);
      Init(clrGrid, [asgDetail2]);

      with Query1, asgDetail do
      begin

        { ���� �÷� ���� }
        RowCount := 50;
        FixedRows := 2;
        FixedCols := 1;
        Cells[FixedCols - 1,0] := ' 7      ��';
        Cells[FixedCols - 1,1] := '��      ��';
        ColWidths[FixedCols - 1] := 100;

        { �÷����� [����, ����] }
          Close;
          SQL.Clear;
          SQL.Add('  SELECT A.DATECD, A.YOILCD, A.DATEGB                        ');
          SQL.Add('    FROM HRDATEMP A                                          ');
          SQL.Add('   WHERE A.DATECD BETWEEN :BFMONTH AND :AFMONTH              ');
          SQL.Add('   ORDER BY A.DATECD                                         ');


          ParamByName('BFMONTH').AsString := cbxBFYear.Text+'-07-01';
          ParamByName('AFMONTH').AsString := cbxBFYear.Text+'-07-31';

          Open;
          { ���� �÷� ���� }
          ColCount := RecordCount + 2;
          MergeCells(ColCount - 1, 0, 1, 2);
          Cells[ColCount - 1,0]  := '��    ��';
          colSeq := FixedCols;
          while not Eof do
          begin
            Cells[colSeq, 0] := IntToStr(DayOf(FieldByName('DATECD').AsDateTime));
            Cells[colSeq, 1] := FormatSettings.ShortDayNames[FieldByName('YOILCD').AsInteger];
            case FieldByName('DATEGB').AsInteger of
              2,3 :
              begin
                Colors[colSeq, 0] := $00DDDBF8;
                Colors[colSeq, 1] := $00DDDBF8;
              end;
              4,5 :
              begin
                Colors[colSeq, 0] := $00CEEFFE;
                Colors[colSeq, 1] := $00CEEFFE;
              end;
            end;
            ColWidths[colSeq] := 35;
            Inc(colSeq);
            Next;
          end;

          { ������ �� ���� }
          Close;
          SQL.Clear;
          SQL.Add('SELECT A.EMPLNUMB, A.KORENAME, C.WKPLNGRP DEPT1, E.WKPLNGRP DEPT2 ');
          SQL.Add('  FROM HREMPLMT A                                           ');
          SQL.Add('  -- ���μ�                                                 ');
          SQL.Add('  LEFT JOIN HRDEPTCD B                                      ');
          SQL.Add('    ON B.DEPTCODE = A.DEPTCODE                              ');
          SQL.Add('   AND B.ENDDDATE = ''9999-99-99''                          ');
          SQL.Add('  LEFT JOIN HRDEPTCD C                                      ');
          SQL.Add('    ON C.DEPTCODE = B.WKPLNGRP                              ');
          SQL.Add('   AND C.ENDDDATE = ''9999-99-99''                          ');
          SQL.Add('  -- ���Ӻμ�                                               ');
          SQL.Add('  LEFT JOIN HRDEPTCD D                                      ');
          SQL.Add('    ON D.DEPTCODE = A.DEPTOVER                              ');
          SQL.Add('   AND D.ENDDDATE = ''9999-99-99''                          ');
          SQL.Add('  LEFT JOIN HRDEPTCD E                                      ');
          SQL.Add('    ON E.DEPTCODE = D.WKPLNGRP                              ');
          SQL.Add('   AND E.ENDDDATE = ''9999-99-99''                          ');
          { �������� ���� ���Ŀ�û }
          if deptCode = '221' then
          begin
            SQL.Add('  LEFT JOIN HRBASECD F                                    ');
            SQL.Add('    ON F.CODEKEY1 = ''64''                                ');
            SQL.Add('   AND F.CODEKEY2 = A.EMPLNUMB                            ');
          end;
          SQL.Add('  LEFT JOIN HRBASECD EP                                     ');  // Ư��������� �߰�(������)
          SQL.Add('    ON EP.CODEKEY1 = ''63''                                 ');
          SQL.Add('   AND EP.CODEKEY2 = A.EMPLNUMB                             ');
          SQL.Add(' WHERE ((A.RETRDATE IS NULL) OR                             ');
          SQL.Add('        (A.RETRDATE IS NOT NULL AND A.RETRDATE > TO_CHAR(SYSDATE, ''YYYY-MM-DD'')))  ');
          SQL.Add('   AND (C.WKPLNGRP = :DEPTCODE OR E.WKPLNGRP = :DEPTCODE)   ');
          SQL.Add('   AND A.JIGWICOD >= ''I7''                                 ');
          SQL.Add('   AND EP.CODEKEY2 IS NULL                                            ');
          { �������� ���� ���Ŀ�û }
          if deptCode = '221' then
            SQL.Add(' ORDER BY TO_NUMBER(F.DSPSEQNO), A.JIGWICOD, NVL(A.TRANDATE, A.EMPLDATE), A.HOBONGCD DESC, A.EMPLNUMB  ')
          else
            SQL.Add(' ORDER BY A.JIGWICOD, NVL(A.TRANDATE, A.EMPLDATE), A.HOBONGCD DESC, A.EMPLNUMB  ');
          { Setting Parameter }
          ParamByName('DEPTCODE').AsString := deptCode;
          Open;

          if Query1.IsEmpty then
            raise Exception.Create(cbxSDept.Text + ' �ް���ȹ�μ��� �������� �ʽ��ϴ�. �μ������� Ȯ�ιٶ��ϴ�.');


          rowSeq := FixedRows;
          while not Eof do
          begin
            Cells[0, rowSeq] := FieldByName('KORENAME').AsString;
            Objects[0, rowSeq] := TMatchStr.Create(FieldByName('EMPLNUMB').AsString);
            Colors[0, rowSeq] := clInfoBk;

            { �����μ��� �����ϴ� ����϶� }
            if FieldByName('DEPT2').AsString <> '' then
            begin
              Query2.Close;
              Query2.SQL.Clear;
              Query2.SQL.Add('SELECT A.WORKDATE, A.INPUTFLG, A.WORKCODE, A.REMARKDT              ');
              Query2.SQL.Add('     , CASE WHEN A.INPUTFLG = ''40'' THEN D.CODECONT END WORKNAME  ');
              Query2.SQL.Add('     , E.DATEGB                                                    ');
              Query2.SQL.Add('  FROM HRWKPLNB A                                                  ');
              Query2.SQL.Add('  JOIN HRWKPLNH B                                                  ');
              Query2.SQL.Add('    ON B.PLNMONTH = A.PLNMONTH                                     ');
              Query2.SQL.Add('   AND B.DEPTCODE = A.DEPTCODE                                     ');
              Query2.SQL.Add('  JOIN HREMPLMT C                                                  ');
              Query2.SQL.Add('    ON C.EMPLNUMB = A.EMPLNUMB                                     ');
              Query2.SQL.Add('  JOIN HRDATEMP E                                                  ');
              Query2.SQL.Add('    ON E.DATECD = A.WORKDATE                                       ');
              Query2.SQL.Add('  LEFT JOIN HRBASECD D                                             ');
              Query2.SQL.Add('    ON D.CODEKEY1 = A.INPUTFLG                                     ');
              Query2.SQL.Add('   AND D.CODEKEY2 = A.WORKCODE                                     ');
              Query2.SQL.Add(' WHERE A.PLNMONTH = :PLNMONTH                                      ');
              Query2.SQL.Add('   AND (A.DEPTCODE = :DEPT1 OR A.DEPTCODE = :DEPT2)                ');
              Query2.SQL.Add('   AND A.EMPLNUMB = :EMPLNUMB                                      ');
              Query2.SQL.Add(' ORDER BY A.WORKDATE                                               ');
              Query2.ParamByName('PLNMONTH').AsString := FormatDateTime('yyyy-mm', afDate);
              Query2.ParamByName('DEPT1').AsString    := Query1.FieldByName('DEPT1').AsString;
              Query2.ParamByName('DEPT2').AsString    := Query1.FieldByName('DEPT2').AsString;
              Query2.ParamByName('EMPLNUMB').AsString := Query1.FieldByName('EMPLNUMB').AsString;
              Query2.Open;

            end;

            Inc(rowSeq);
            Next;

         end;

         if deptCode = '221' then    //�ð��� ���� �����.... ������� 8���̳Ѵ� �μ� �ϵ��ڵ�. ���� ����� ���ؼ� if�� ������
          RowCount := rowSeq
         else
          RowCount := 10;

      end;


      with Query1, asgDetail2 do
      begin
        { ���� �÷� ���� }
        RowCount := 50;
        FixedRows := 2;
        FixedCols := 1;
        Cells[FixedCols - 1,0] := ' 8      ��';
        Cells[FixedCols - 1,1] := '��      ��';
        ColWidths[FixedCols - 1] := 100;


        { �÷����� [����, ����] }
          Close;
          SQL.Clear;
          SQL.Add('  SELECT A.DATECD, A.YOILCD, A.DATEGB                        ');
          SQL.Add('    FROM HRDATEMP A                                          ');
          SQL.Add('   WHERE A.DATECD BETWEEN :BFMONTH AND :AFMONTH              ');
          SQL.Add('   ORDER BY A.DATECD                                         ');

          ParamByName('BFMONTH').AsString := cbxBFYear.Text+'-08-01';
          ParamByName('AFMONTH').AsString := cbxBFYear.Text+'-08-31';

          Open;
          { ���� �÷� ���� }
          ColCount := RecordCount + 2;
          MergeCells(ColCount - 1, 0, 1, 2);
          Cells[ColCount - 1,0]  := '��    ��';
          colSeq := FixedCols;
          while not Eof do
          begin
            Cells[colSeq, 0] := IntToStr(DayOf(FieldByName('DATECD').AsDateTime));
            Cells[colSeq, 1] := FormatSettings.ShortDayNames[FieldByName('YOILCD').AsInteger];
            case FieldByName('DATEGB').AsInteger of
              2,3 :
              begin
                Colors[colSeq, 0] := $00DDDBF8;
                Colors[colSeq, 1] := $00DDDBF8;
              end;
              4,5 :
              begin
                Colors[colSeq, 0] := $00CEEFFE;
                Colors[colSeq, 1] := $00CEEFFE;
              end;
            end;
            ColWidths[colSeq] := 35;
            Inc(colSeq);
            Next;
          end;

          { ������ �� ���� }
          Close;
          SQL.Clear;
          SQL.Add('SELECT A.EMPLNUMB, A.KORENAME, C.WKPLNGRP DEPT1, E.WKPLNGRP DEPT2 ');
          SQL.Add('  FROM HREMPLMT A                                           ');
          SQL.Add('  -- ���μ�                                                 ');
          SQL.Add('  LEFT JOIN HRDEPTCD B                                      ');
          SQL.Add('    ON B.DEPTCODE = A.DEPTCODE                              ');
          SQL.Add('   AND B.ENDDDATE = ''9999-99-99''                          ');
          SQL.Add('  LEFT JOIN HRDEPTCD C                                      ');
          SQL.Add('    ON C.DEPTCODE = B.WKPLNGRP                              ');
          SQL.Add('   AND C.ENDDDATE = ''9999-99-99''                          ');
          SQL.Add('  -- ���Ӻμ�                                               ');
          SQL.Add('  LEFT JOIN HRDEPTCD D                                      ');
          SQL.Add('    ON D.DEPTCODE = A.DEPTOVER                              ');
          SQL.Add('   AND D.ENDDDATE = ''9999-99-99''                          ');
          SQL.Add('  LEFT JOIN HRDEPTCD E                                      ');
          SQL.Add('    ON E.DEPTCODE = D.WKPLNGRP                              ');
          SQL.Add('   AND E.ENDDDATE = ''9999-99-99''                          ');
          { �������� ���� ���Ŀ�û }
          if deptCode = '221' then
          begin
            SQL.Add('  LEFT JOIN HRBASECD F                                    ');
            SQL.Add('    ON F.CODEKEY1 = ''64''                                ');
            SQL.Add('   AND F.CODEKEY2 = A.EMPLNUMB                            ');
          end;
          SQL.Add('  LEFT JOIN HRBASECD EP                                     ');  // Ư��������� �߰�(������)
          SQL.Add('    ON EP.CODEKEY1 = ''63''                                 ');
          SQL.Add('   AND EP.CODEKEY2 = A.EMPLNUMB                             ');
          SQL.Add(' WHERE ((A.RETRDATE IS NULL) OR                             ');
          SQL.Add('        (A.RETRDATE IS NOT NULL AND A.RETRDATE > TO_CHAR(SYSDATE, ''YYYY-MM-DD'')))  ');
          SQL.Add('   AND (C.WKPLNGRP = :DEPTCODE OR E.WKPLNGRP = :DEPTCODE)   ');
          SQL.Add('   AND A.JIGWICOD >= ''I7''                                 ');
          SQL.Add('   AND EP.CODEKEY2 IS NULL                                            ');
          { �������� ���� ���Ŀ�û }
          if deptCode = '221' then
            SQL.Add(' ORDER BY TO_NUMBER(F.DSPSEQNO), A.JIGWICOD, NVL(A.TRANDATE, A.EMPLDATE), A.HOBONGCD DESC, A.EMPLNUMB  ')
          else
            SQL.Add(' ORDER BY A.JIGWICOD, NVL(A.TRANDATE, A.EMPLDATE), A.HOBONGCD DESC, A.EMPLNUMB  ');
          { Setting Parameter }
          ParamByName('DEPTCODE').AsString := deptCode;
          Open;

          if Query1.IsEmpty then
            raise Exception.Create(cbxSDept.Text + ' �ް���ȹ�μ��� �������� �ʽ��ϴ�. �μ������� Ȯ�ιٶ��ϴ�.');


          rowSeq := FixedRows;
          while not Eof do
          begin
            Cells[0, rowSeq] := FieldByName('KORENAME').AsString;
            Objects[0, rowSeq] := TMatchStr.Create(FieldByName('EMPLNUMB').AsString);
            Colors[0, rowSeq] := clInfoBk;

            { �����μ��� �����ϴ� ����϶� }
            if FieldByName('DEPT2').AsString <> '' then
            begin
              Query2.Close;
              Query2.SQL.Clear;
              Query2.SQL.Add('SELECT A.WORKDATE, A.INPUTFLG, A.WORKCODE, A.REMARKDT              ');
              Query2.SQL.Add('     , CASE WHEN A.INPUTFLG = ''40'' THEN D.CODECONT END WORKNAME  ');
              Query2.SQL.Add('     , E.DATEGB                                                    ');
              Query2.SQL.Add('  FROM HRWKPLNB A                                                  ');
              Query2.SQL.Add('  JOIN HRWKPLNH B                                                  ');
              Query2.SQL.Add('    ON B.PLNMONTH = A.PLNMONTH                                     ');
              Query2.SQL.Add('   AND B.DEPTCODE = A.DEPTCODE                                     ');
              Query2.SQL.Add('  JOIN HREMPLMT C                                                  ');
              Query2.SQL.Add('    ON C.EMPLNUMB = A.EMPLNUMB                                     ');
              Query2.SQL.Add('  JOIN HRDATEMP E                                                  ');
              Query2.SQL.Add('    ON E.DATECD = A.WORKDATE                                       ');
              Query2.SQL.Add('  LEFT JOIN HRBASECD D                                             ');
              Query2.SQL.Add('    ON D.CODEKEY1 = A.INPUTFLG                                     ');
              Query2.SQL.Add('   AND D.CODEKEY2 = A.WORKCODE                                     ');
              Query2.SQL.Add(' WHERE A.PLNMONTH = :PLNMONTH                                      ');
              Query2.SQL.Add('   AND (A.DEPTCODE = :DEPT1 OR A.DEPTCODE = :DEPT2)                ');
              Query2.SQL.Add('   AND A.EMPLNUMB = :EMPLNUMB                                      ');
              Query2.SQL.Add(' ORDER BY A.WORKDATE                                               ');
              Query2.ParamByName('PLNMONTH').AsString := FormatDateTime('yyyy-mm', afDate);
              Query2.ParamByName('DEPT1').AsString    := Query1.FieldByName('DEPT1').AsString;
              Query2.ParamByName('DEPT2').AsString    := Query1.FieldByName('DEPT2').AsString;
              Query2.ParamByName('EMPLNUMB').AsString := Query1.FieldByName('EMPLNUMB').AsString;
              Query2.Open;

            end;

            Inc(rowSeq);
            Next;

         end;

         if deptCode = '221' then     //�ð��� ���� �����.... ������� 8���̳Ѵ� �μ� �ϵ��ڵ�. ���� ����� ���ؼ� if�� ������
          RowCount := rowSeq
         else
          RowCount := 10;
      end;



      asgDetail.SetFocus;
    except
      on E: Exception do
      begin
        ModeControl(mcNone, canModWkPln);
        CALL_ERROR('��ȸ ����', fmTITLE.Caption + '��ȸ �� ������ �߻��߽��ϴ�.' + #13#10#13#10 + E.Message);
        Exit;
      end;
    end;
  finally
    asgDetail.EndUpdate;
    asgDetail2.EndUpdate;
  end;
end;

procedure THB40010.ModeControl(mode: TEModeControl; appSts:Boolean = True);
begin

      pnlPrepareDetailGrid.Visible := False;

end;

//procedure THB40010.sbAddNewWorkPlanClick(Sender: TObject);
//begin
//  inherited;
//  { �α��� ���� �ٹ���ȹǥ �׷� �μ��� �ڵ� ���� }
//  cbxDept.ItemIndex := cbxDept.Items.IndexOf(FUserDeptName);
//
//  if IsDate(cbxPlanYear.Text + '-' + cbxPlanMonth.Text + '-' + '1') then
//    DetailQuery(Sender);
//end;

//procedure THB40010.sbCopyWPPreMonthClick(Sender: TObject);
//var
//  plnMonth:TDateTime;
//begin
//  try
//    try
//      { ��ȿ���˻� }
//      if asgDetail.Cells[0, asgDetail.FixedRows] = EmptyStr then
//        raise Exception.Create('�ٹ���ȹǥ�� �������� �ʾҽ��ϴ�.');
//
//      if not IsDate(cbxPlanYear.Text + '-' + cbxPlanMonth.Text + '-' + '01') then
//        raise Exception.Create('[��ȹ���] �׸��� ���������� �Է���� Ȯ�ιٶ��ϴ�.');
//
//      if cbxDept.ItemIndex = -1 then
//        raise Exception.Create('[�μ�] �׸��� ���������� �Է���� Ȯ�ιٶ��ϴ�.');
//
//      plnMonth := IncMonth(EncodeDate(StrToInt(cbxPlanYear.Text), StrToInt(cbxPlanMonth.Text), 1), -1);
//      with Query1 do
//      begin
//        Close;
//        SQL.Clear;
//        SQL.Add('SELECT ISSUENUM, MEMOREF1    ');
//        SQL.Add('  FROM HRWKPLNH              ');
//        SQL.Add(' WHERE PLNMONTH = :PLNMONTH  ');
//        SQL.Add('   AND DEPTCODE = :DEPTCODE  ');
//        ParamByName('PLNMONTH').AsString := FormatDateTime('yyyy-mm', plnMonth);
//        ParamByName('DEPTCODE').AsString := TMatchStr(cbxDept.Items.Objects[cbxDept.ItemIndex]).FMatchStr;
//        Open;
//
//        if Query1.IsEmpty then
//          raise Exception.Create(FormatDateTime('yyyy-mm', plnMonth) +'�� ' + cbxDept.Text + ' �� �ٹ���ȹǥ�� ��ϵǾ� ���� �ʽ��ϴ�.');
//
//        edtIssueNum.Text    := FieldByName('ISSUENUM').AsString;
//        memoRef1.Lines.Text := FieldByName('MEMOREF1').AsString;
////        memoRef2.Lines.Text := FieldByName('MEMOREF2').AsString;
//      end;
//    except
//      on E: Exception do
//      begin
//        CALL_ERROR('���� ����', '���� ������� ���� �� ���� �߻��߽��ϴ�.' + #13#10#13#10 + E.Message);
//      end;
//    end;
//  finally
//  end;
//end;

procedure THB40010.sbDetailInitClick(Sender: TObject);
var
  r,c:Integer;
begin
  inherited;
  with asgDetail do
  begin
    if IsDate(cbxBFYear.Text + '-' + '07' + '-' + '1') and
      (cbxSDept.ItemIndex > -1) then
    begin
      { mode �� ���� detail grid ���� }
      if Length(Cells[0, FixedRows]) > 0 then
        ClearRect(1, FixedRows, ColCount - 1, RowCount - 1 );
    end;
  end;
end;



procedure THB40010.asgDetailGetAlignment(Sender: TObject; ARow, ACol: Integer;
  var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  inherited;
  if asgDetail.IsFixed(ACol, ARow) then
    HAlign := taCenter
  else
  if asgDetail.ColCount - 1 = ACol then
    HAlign := taLeftJustify
  else if asgDetail2.IsFixed(ACol, ARow) then
    HAlign := taCenter
  else
  if asgDetail2.ColCount - 1 = ACol then
    HAlign := taLeftJustify
  else
    HAlign := taCenter;
end;

procedure THB40010.asgDetailGetCellColor(Sender: TObject; ARow, ACol: Integer;
  AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
begin
  inherited;
  if ARow < asgDetail.FixedRows then
  begin
    if (ARow = 0) and (asgDetail.Cells[ACol, ARow] = 'DOCK') then
      AFont.Size := 8;
  end
  else
  begin
    { �ָ� �Ǵ� ������ �� ��Ÿ�� ��û }
    if asgDetail.Colors[ACol, 0] <> clNone then
    begin
      AFont.Color := clRed;
      AFont.Style := [fsBold];
    end
    else if asgDetail2.Colors[ACol, 0] <> clNone then
    begin
      AFont.Color := clRed;
      AFont.Style := [fsBold];
    end;

  end;
end;


procedure THB40010.cbxConDateChange(Sender: TObject);
var
  bfYear, afYear, bfMonth, afMonth : Word;
  bfDate, afDate:TDateTime;
begin
  inherited;
    pnlPrepareDetailGrid.Visible := true;
    asgDetail.Visible := false;
    asgDetail2.Visible := false;
    pnl1.Visible := false;
    panel5.Visible := false;
  { ���� 4����� ��� ���õ������ }
  if (cbxBFYear.ItemIndex > -1) then
  begin
    { �׸��忡 ������ ���� �� �ʱ�ȭ }
    if Length(asgMaster.Cells[0, asgMaster.FixedRows]) > 0 then
      SB_QUERYClick(nil);

    { �ʱ�ȭ }
    if Length(asgDetail.Cells[0, asgMaster.FixedRows]) > 0 then
    begin
      Init(clrGrid, [asgDetail]);
       Init(clrGrid, [asgDetail2]);
      pnlPrepareDetailGrid.Visible := True;
    end;

    bfYear := StrToInt(cbxBFYear.Text);

    { �񱳳�¥���� }
    bfDate := EncodeDate(bfYear, 07, 1);

    with Query1 do
    begin
      Close;
      SQL.Clear;

      SQL.Add('  SELECT DEPTNAME, DEPTCODE, STATDATE, ENDDDATE                                                                ');
      SQL.Add('    FROM HRDEPTCD                                                                                              ');
      SQL.Add('   WHERE SUBSTR(ENDDDATE, 1, 7) > :ENDDATE1                                                                    ');
      SQL.Add('   AND DEPTCODE IN (''090'', ''610'', ''310'', ''222'', ''210'', ''240'', ''221'', ''400'')                    ');

      ParamByName('ENDDATE1').AsString := FormatDateTime('yyyy', bfDate)+'-07';
      Open;

      cbxSDept.Clear;
      //cbxSDept.Items.Add('��ü�μ�');
      while not Eof do
      begin
        cbxSDept.AddItem(Fields[0].AsString, TMatchStr.Create(Fields[1].AsString));
       Next;
      end;

      cbxSDept.ItemIndex := -1;
    end;
  end;
end;

//procedure THB40010.cbxSDeptChange(Sender: TObject);
//begin
//  inherited;
//  { �׸��忡 ������ ���� �� ����ȸ }
//  if Length(asgMaster.Cells[0, asgMaster.FixedRows]) > 0 then
//    SB_QUERYClick(nil);
//end;

procedure THB40010.ChangeEnrollInfo(Sender: TObject);
var
  date1, date2:TDateTime;
begin
  inherited;
  if IsDate(cbxBFYear.Text + '-' + '07' + '-' + '1') then
  begin
    date1 := StrToDate(cbxBFYear.Text + '-' + '07' + '-' + '1');
    date2 := StrToDate(cbxBFYear.Text + '-' + '08' + '-' + '1');
    lblTitleDetailGrid.Caption :=
    '������ ( ' + FormatDateTime('yyyy-mm',date1) +'~' + FormatDateTime('yyyy-mm',date2) + ' )'
  end
  else
    lblTitleDetailGrid.Caption := '������';
  if IsDate(cbxBFYear.Text + '-' + '07' + '-' + '1') and
     (cbxSDept.ItemIndex > -1) then
    DetailQuery(Sender);
end;

{ TMatchStr }
constructor TMatchStr.Create(matchStr: string);
begin
  FMatchStr := matchStr;
end;

procedure THB40010.asgDetailGetEditorType(Sender: TObject; ACol, ARow: Integer;
  var AEditor: TEditorType);
begin
  inherited;
  if asgDetail.ColCount - 1 > ACol then
  begin
    AEditor := edUpperCase;
    asgDetail.MaxEditLength := 1;
  end
  else if asgDetail2.ColCount - 1 > ACol then
  begin
    AEditor := edUpperCase;
    asgDetail.MaxEditLength := 1;
  end
  else
    asgDetail.MaxEditLength := 0;
end;

procedure THB40010.asgDetailKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  i:Integer;
  gc:TGridCoord;
  comment:string;
begin
  inherited;
  if Key = VK_DELETE then
  begin
    with asgDetail do
    begin
      for i := 0 to SelectedCellsCount - 1 do
      begin
        gc := SelectedCell[i];
        Cells[gc.X, gc.Y] := '';
        if IsComment(gc.X, gc.Y, comment) then
          RemoveComment(gc.X, gc.Y);
      end;
    end;
    with asgDetail2 do
    begin
      for i := 0 to SelectedCellsCount - 1 do
      begin
        gc := SelectedCell[i];
        Cells[gc.X, gc.Y] := '';
        if IsComment(gc.X, gc.Y, comment) then
          RemoveComment(gc.X, gc.Y);
      end;
    end;

  end;
end;

procedure THB40010.asgDetailKeyPress(Sender: TObject; var Key: Char);
begin
  inherited;
  if asgDetail.ColCount - 1 > asgDetail.Col then
  begin
    if WorkTypeShortCut.ContainsKey(Key) then
      Key := WorkTypeShortCut.Items[Key];
  end
  else if asgDetail.ColCount - 1 > asgDetail.Col then
    begin
    if WorkTypeShortCut.ContainsKey(Key) then
      Key := WorkTypeShortCut.Items[Key];
  end;

end;


procedure THB40010.btnConfirmShortCutClick(Sender: TObject);
begin
  inherited;
  Application.MessageBox(PChar(HintShorCutList.Text),'�Է´���Ű �������', MB_OK or MB_ICONINFORMATION)
end;

end.
