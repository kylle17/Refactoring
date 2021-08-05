package query;

//기본 Java API
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Iterator;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import common.CommonLib;
import common.DBManager;
import common.LoggableStatement;
import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

public class kdbp_excel_poi_query {
	
	CommonLib util = new CommonLib();      // 공통모듈호출
	
	db_conn 	db_conn	= new db_conn();  
	String dbConnName  = db_conn.db_conn();
	
	String dbConnOpenMsg  = "Con 연결: kdbp_excel_query.";
	String dbConnCloseMsg = "Con 해제: kdbp_excel_query.";
	
	/***********************************************
     * getSysDat
     * description : 오라클 시간 얻기
     * table       : 
     * param       : 
     * return      : 
    /************************************************/
	public String getSysDat(Connection conn, String format) {

		ResultSet            rs = null;
 		PreparedStatement pstmt = null;
 		StringBuffer    sql_str = new StringBuffer();
 		String result = "";
 		String sysdat= "";

     	try {
 			// 풀을 통해서 컨넥션을 얻음.
     		DBManager.getConn("연결: kdbp_excel_query.getSysDat()");

 			// SQL문장 생성
 			sql_str.append(" select to_char(sysdate, 'YYYY-MM-DD HH:MI:SS') from dual    \n");
	
 			pstmt = new LoggableStatement(conn, sql_str.toString());
 			System.out.println(((LoggableStatement)pstmt).getQueryString());
 			rs = pstmt.executeQuery();
 	

 			if (rs.next()) {
 				sysdat = util.checkNull(rs.getString(1)); //2013-07-03-08.32.59.755985
 				
 				if(format.equals("yyyy-MM-dd")){
 					result = sysdat.substring(0, 10);
 				}else if(format.equals("yyyy-MM-dd HH:mm")){
 					result = sysdat.substring(0, 10)+" "+sysdat.substring(11, 13)+":"+sysdat.substring(14, 16);
 				}else if(format.equals("yyyy-MM-dd HH:mm:ss")){
 					result = sysdat.substring(0, 10)+" "+sysdat.substring(11, 13)+":"+sysdat.substring(14, 16)+":"+sysdat.substring(17, 19);
 				}else if(format.equals("yyyy-MM-dd HH:mm:ss.SSS")){
 					result = sysdat.substring(0, 10)+" "+sysdat.substring(11, 13)+":"+sysdat.substring(14, 16)+":"+sysdat.substring(17, 19)+"."+sysdat.substring(20, 23);
 				}
 			} 

 			
     	}catch(SQLException se){
  			se.printStackTrace();

  		}catch(Exception e){
  			e.printStackTrace();

  		}finally{
     	    DBManager.close("해제: kdbp_excel_query.getSysDat()", rs, pstmt);
         }

 		return result;
     }
	
	/***********************************************
     * getSysDat
     * description : 벨류값 타입별로 지정해서 가져오기
     * table       : 
     * param       : 
     * return      : 
    /************************************************/
	public String matchCellType_xls(HSSFCell cell) {
        String value="";
        try {
        	//셀이 빈값일경우를 위한 널체크
        	//타입별로 내용 읽기
        	switch (cell.getCellType()){
        	case HSSFCell.CELL_TYPE_FORMULA:
        		switch(cell.getCachedFormulaResultType()) {
                case HSSFCell.CELL_TYPE_NUMERIC:
                	cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                	value = cell.getStringCellValue();
                    break;
                case HSSFCell.CELL_TYPE_STRING:
                	value = cell.getStringCellValue();
                    break;
        		}
        		break;
        	case HSSFCell.CELL_TYPE_NUMERIC:     
        		value= String.valueOf( cell.getNumericCellValue() )+"";
        		break;
        	case HSSFCell.CELL_TYPE_STRING:
        		value=cell.getStringCellValue().replaceAll("<", "＜").replaceAll(">", "＞")+"";
        		break;
        	case HSSFCell.CELL_TYPE_BOOLEAN:
        		value=cell.getBooleanCellValue()+"";
        		break;
        	case HSSFCell.CELL_TYPE_BLANK:
        		value="";
        		break;
        	case HSSFCell.CELL_TYPE_ERROR:
        		value="";
        		break;
        	default:
        		value="";
        		break;
        	}
			
		} catch (Exception e) {
			//System.out.println("매치타입 오류 "+e);
		}
        return value;
     }
	public String matchCellType_xlsx(XSSFCell cell) {
		String value="";
		try {
			
			
			
			//셀이 빈값일경우를 위한 널체크
			//타입별로 내용 읽기
			switch (cell.getCellType()){
			case XSSFCell.CELL_TYPE_FORMULA:
        		switch(cell.getCachedFormulaResultType()) {
                case XSSFCell.CELL_TYPE_NUMERIC:
                	cell.setCellType(XSSFCell.CELL_TYPE_STRING);
                	value = cell.getStringCellValue();
                    break;
                case XSSFCell.CELL_TYPE_STRING:
                	value = cell.getStringCellValue();
                    break;
        		}
        		break;
			case XSSFCell.CELL_TYPE_NUMERIC:
				value= String.valueOf(cell.getNumericCellValue())+"";
				break;
			case XSSFCell.CELL_TYPE_STRING:
				value=cell.getStringCellValue().replaceAll("<", "＜").replaceAll(">", "＞")+"";
				break;
	        case HSSFCell.CELL_TYPE_BOOLEAN:
		        value=cell.getBooleanCellValue()+"";
		        break;
			case XSSFCell.CELL_TYPE_BLANK:
				value="";
				break;
			case XSSFCell.CELL_TYPE_ERROR:
				value="";
				break;
			default:
				value="";
				break;
			}
		} catch (Exception e) {
			//System.out.println("매치타입 오류 "+e);
		} 
		return value;
	}
	
	
	/***********************************************
     * getSysDat
     * description : 벨류값 타입별로 지정해서 가져오기
     * table       : 
     * param       : 
     * return      : 
    /************************************************/
	public String matchCellType(Cell cell) {
		String value = "";
        try {
        	//셀이 빈값일경우를 위한 널체크
        	//타입별로 내용 읽기
        	switch (cell.getCellType()){
        	case HSSFCell.CELL_TYPE_FORMULA:
        		switch(cell.getCachedFormulaResultType()) {
                case HSSFCell.CELL_TYPE_NUMERIC:
                	value = cell.getNumericCellValue()+"";
                    break;
                case HSSFCell.CELL_TYPE_STRING:
                	value = cell.getStringCellValue();
                    break;
        		}
        		break;
        	case HSSFCell.CELL_TYPE_NUMERIC:
        		value=cell.getNumericCellValue()+"";
        		break;
        	case HSSFCell.CELL_TYPE_STRING:
        		value=cell.getStringCellValue().replaceAll("<", "＜").replaceAll(">", "＞")+"";
        		break;
        	case HSSFCell.CELL_TYPE_BOOLEAN:
        		value=cell.getBooleanCellValue()+"";
        		break;
        	case HSSFCell.CELL_TYPE_BLANK:
        		value="";
        		break;
        	case HSSFCell.CELL_TYPE_ERROR:
        		value="";
        		break;
        	default:
        		value="";
        		break;
        	}
			
    		switch (cell.getCellType()){
    		case XSSFCell.CELL_TYPE_FORMULA:
        		switch(cell.getCachedFormulaResultType()) {
                case XSSFCell.CELL_TYPE_NUMERIC:
                	value = cell.getNumericCellValue()+"";
                    break;
                case XSSFCell.CELL_TYPE_STRING:
                	value = cell.getStringCellValue();
                    break;
        		}
        		break;
    		case XSSFCell.CELL_TYPE_NUMERIC:
    			value=cell.getNumericCellValue()+"";
    			break;
    		case XSSFCell.CELL_TYPE_STRING:
    			value=cell.getStringCellValue().replaceAll("<", "＜").replaceAll(">", "＞")+"";
    			break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
    	        value=cell.getBooleanCellValue()+"";
    	        break;
    		case XSSFCell.CELL_TYPE_BLANK:
    			value="";
    			break;
    		case XSSFCell.CELL_TYPE_ERROR:
    			value="";
    			break;
    		default:
    			value="";
    			break;
    		}
        	
		} catch (Exception e) {
			//System.out.println("매치타입 오류 "+e);
		}

		return value;
	}
	
	

	public JSONObject XLStoJSONObject_FIXAMT(String Jobgbn,String ODDGBN,String ODDGBNNM,File file, String USERID) throws Exception{
		
		System.out.println("file : "+file);
		boolean flag = true;
		JSONObject 	result 	= new JSONObject();
		FileInputStream fis=new FileInputStream(file);
		String msg = "처리되었습니다.\n자료를 확인 후 저장해주세요.";
		
		String extension = FilenameUtils.getExtension(file.getName());
		System.out.println("확장자 : "+extension);
		
		try {
		
			// 엑셀 내용을 읽어서 담을 객체의 참조변수 준비. 
			org.apache.poi.ss.usermodel.Workbook workbook = null; // 엑셀 파일에 해당하는 객체 담을 참조변수 선언.
			org.apache.poi.ss.usermodel.Sheet sheet = null; // 엑셀 시트에 해당하는 객체 담을 참조변수 선언. 
			org.apache.poi.ss.usermodel.Row row = null; // 엑셀 row에 해당하는 객체 생성.
			
			// 다형성 활용 객체 생성하기. 
			if(extension.equals("xls")) {	// 확장자 xls인 경우 HSSFWorkbook객체 생성
				workbook = new HSSFWorkbook(fis); // 액셀 파일에 해당하는 객체 생성. 
				sheet = (HSSFSheet)workbook.getSheetAt(0); 	// 0번째 시트의 내용을 가져와서 객체에 담는다. 
			}else {							// 확장자 xlsx인 경우 XSSFWorkbook객체 생성
				workbook = new XSSFWorkbook(fis); // 엑셀 파일에 해당하는 객체 생성. 
				sheet = (XSSFSheet)workbook.getSheetAt(0); 	// 0번째 시트의 내용을 가져와서 객체에 담는다.
			};
		
				if( sheet != null)	{
					try {
				
						
						// 일반 로직 (FIXAMT_VIEW)
						int nRowEndIndex   = sheet.getPhysicalNumberOfRows();
						int ArrCnt = 0;						
						JSONObject resobj1 = new JSONObject();
						JSONArray  FIXAMT_VIEW_arr = new JSONArray();
						System.out.println("ODDGBN: "+ODDGBN);
						System.out.println(nRowEndIndex);
						for(int rowindex=2; rowindex<nRowEndIndex; rowindex++){
							row=sheet.getRow(rowindex);
							if( row != null)	{
								String ACUNIT = util.checkNull(matchCellType(row.getCell( 0)));
								String AREANM = util.checkNull(matchCellType(row.getCell( 1)));
								String ACCNUM = util.checkNull(matchCellType(row.getCell( 3)));
								String ACDATE = util.checkNull(matchCellType(row.getCell( 6)));
								String GYPBUS = util.checkNull(matchCellType(row.getCell( 7)));
								String GYPMAN = util.checkNull(matchCellType(row.getCell( 8)));
								String ACDEXT = util.checkNull(matchCellType(row.getCell( 9)));
								String ACDNAM = util.checkNull(matchCellType(row.getCell(10)));
								String REMARK = util.checkNull(matchCellType(row.getCell(11)));
								String YESBUS = util.checkNull(matchCellType(row.getCell(13)));
								String SPCBUS = util.checkNull(matchCellType(row.getCell(14)));
								String BUSNAM = util.checkNull(matchCellType(row.getCell(15)));
								String CHAAMT = util.checkNull(matchCellType(row.getCell(19)));
								String DAEAMT = util.checkNull(matchCellType(row.getCell(20)));
								resobj1.put("ACUNIT", ACUNIT);
								resobj1.put("AREANM", AREANM);
								resobj1.put("ACCNUM", ACCNUM);
								resobj1.put("ACDATE", ACDATE);
								resobj1.put("GYPBUS", GYPBUS);
								resobj1.put("GYPMAN", GYPMAN);
								resobj1.put("ACDEXT", ACDEXT);
								resobj1.put("ACDNAM", ACDNAM);
								resobj1.put("REMARK", REMARK);
								resobj1.put("YESBUS", YESBUS);
								resobj1.put("SPCBUS", SPCBUS);
								resobj1.put("BUSNAM", BUSNAM);
								resobj1.put("CHAAMT", CHAAMT);
								resobj1.put("DAEAMT", DAEAMT);
								resobj1.put("ODDGBN", ODDGBN);
								resobj1.put("ODDGBNNM", ODDGBNNM);
								FIXAMT_VIEW_arr.add(resobj1);
							}
							
						}

						result.put("rs1", FIXAMT_VIEW_arr);
						result.put("rsCount", ArrCnt);
						
						
						// 합산로직 (FIXAMT)
						nRowEndIndex   = sheet.getPhysicalNumberOfRows();
						ArrCnt = 0;
						JSONObject resobj2 = new JSONObject();
						JSONObject resarr_value_obj = new JSONObject();
						JSONObject resarr_key_obj = new JSONObject();
						JSONArray  resarr = new JSONArray();
						System.out.println("ODDGBN: "+ODDGBN);
						System.out.println(nRowEndIndex);
						for(int rowindex=2; rowindex<nRowEndIndex; rowindex++){
							row=sheet.getRow(rowindex);
							if( row != null)	{
								String ACUNIT = util.checkNull(matchCellType(row.getCell( 0)));
								String AREANM = util.checkNull(matchCellType(row.getCell( 1)));
								String ACCNUM = util.checkNull(matchCellType(row.getCell( 3)));
								String ACDATE = util.checkNull(matchCellType(row.getCell( 6)));
								String GYPBUS = util.checkNull(matchCellType(row.getCell( 7)));
								String GYPMAN = util.checkNull(matchCellType(row.getCell( 8)));
								String ACDEXT = util.checkNull(matchCellType(row.getCell( 9)));
								String ACDNAM = util.checkNull(matchCellType(row.getCell(10)));
								String REMARK = util.checkNull(matchCellType(row.getCell(11)));
								String YESBUS = util.checkNull(matchCellType(row.getCell(13)));
								String SPCBUS = util.checkNull(matchCellType(row.getCell(14)));
								String BUSNAM = util.checkNull(matchCellType(row.getCell(15)));
								String CHAAMT = util.checkNull(matchCellType(row.getCell(19)));
								String DAEAMT = util.checkNull(matchCellType(row.getCell(20)));
								double CHAAMT_temp = 0;
								double DAEAMT_temp = 0;
								resobj2.clear();
								ArrCnt++;
								//resarr_value_obj.keys();
								if(!resarr_value_obj.getJSONObject(ACCNUM+ACDATE+GYPBUS+ACDEXT+SPCBUS+BUSNAM).isEmpty()) {
									ArrCnt--;
									CHAAMT_temp = resarr_value_obj.getJSONObject(ACCNUM+ACDATE+GYPBUS+ACDEXT+SPCBUS+BUSNAM).getDouble("CHAAMT");
									DAEAMT_temp = resarr_value_obj.getJSONObject(ACCNUM+ACDATE+GYPBUS+ACDEXT+SPCBUS+BUSNAM).getDouble("DAEAMT");
								}else {
									resarr_key_obj.put(ArrCnt, ACCNUM+ACDATE+GYPBUS+ACDEXT+SPCBUS+BUSNAM);
								}
								resobj2.put("ACUNIT", ACUNIT);
								resobj2.put("AREANM", AREANM);
								resobj2.put("ACCNUM", ACCNUM);
								resobj2.put("ACDATE", ACDATE);
								resobj2.put("GYPBUS", GYPBUS);
								resobj2.put("GYPMAN", GYPMAN);
								resobj2.put("ACDEXT", ACDEXT);
								resobj2.put("ACDNAM", ACDNAM);
								resobj2.put("REMARK", REMARK);
								resobj2.put("YESBUS", YESBUS);
								resobj2.put("SPCBUS", SPCBUS);
								resobj2.put("BUSNAM", BUSNAM);
								resobj2.put("CHAAMT", Double.parseDouble(CHAAMT.isEmpty()?"0":CHAAMT)+CHAAMT_temp);
								resobj2.put("DAEAMT", Double.parseDouble(DAEAMT.isEmpty()?"0":DAEAMT)+DAEAMT_temp);
								resobj2.put("ODDGBN", ODDGBN);
								resobj2.put("ODDGBNNM", ODDGBNNM);
								resarr_value_obj.put(ACCNUM+ACDATE+GYPBUS+ACDEXT+SPCBUS+BUSNAM, resobj2);
							}
							
						}
						for(int i=1; i<=ArrCnt; i++) {
							resarr.add(resarr_value_obj.get(resarr_key_obj.get(Integer.toString(i))));
						}
						result.put("rs2", resarr);
						result.put("rsCount", ArrCnt);
						
						
					} finally {
						if(workbook!=null){
							workbook.close();
						}
					}
				}else{
					flag = false;
					msg = "엑셀 파일을 읽을수 없습니다.";
				}
		} catch (Exception e) {
			flag = false;
			System.out.println(e);
			msg = "엑셀 파일을 읽는중 에러가 발생하였습니다.";
		} finally {
			result.put("rescod", String.valueOf(flag));
			result.put("resmsg", msg);
		}
		return result;
	}
	
	
	
	
	
}