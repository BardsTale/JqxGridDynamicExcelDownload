package exportExcel.common;

import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;

public class RenderExcelFromJqxGrid {
	@SuppressWarnings({"deprecation","unchecked"})
	public static SXSSFWorkbook makeWorkbook(Map<String, Object> model){
		long t1 = System.currentTimeMillis();
		long t2 = 0;
		SXSSFWorkbook wb = new SXSSFWorkbook(100); // 100개의 row씩을 메모리에 담고 초과 시 서버의 디스크로 플러쉬한다.
		try {
			HashMap<String, Object> param = (HashMap<String, Object>)model.get("model");
			List<HashMap<String, Object>> excelArray = (List<HashMap<String,Object>>)param.get("rows");  //그리드 바디
			List<HashMap<String, Object>> head_excelArray = (List<HashMap<String,Object>>)param.get("head_rows");  //그리드 헤드
 			int headcnt = head_excelArray.size(); //헤드 row 수
 			double headlen = head_excelArray.get(0).size(); //헤드 Column 수
 			int rowcnt = excelArray.size(); //바디 row 수
 			double rowlen = excelArray.get(0).size(); //바디 Column 수, 일반적으로 헤드와 같지만 가로 페이징일 경우 필요.
 			int sheet_cnt = (int) Math.ceil(rowlen/headlen); //가로 페이징 시 시트 수
			
			//column 헤더 스타일
			XSSFCellStyle cellHeadStyle = (XSSFCellStyle) wb.createCellStyle();
			cellHeadStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(248,248,248)));          
			cellHeadStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			cellHeadStyle.setBorderTop(BorderStyle.THIN);
			cellHeadStyle.setTopBorderColor(new XSSFColor(new java.awt.Color(221,221,221))); 
			cellHeadStyle.setBorderBottom(BorderStyle.THIN);
			cellHeadStyle.setBottomBorderColor(new XSSFColor(new java.awt.Color(221,221,221))); 
			cellHeadStyle.setBorderLeft(BorderStyle.THIN);
			cellHeadStyle.setLeftBorderColor(new XSSFColor(new java.awt.Color(221,221,221))); 
			cellHeadStyle.setBorderRight(BorderStyle.THIN);
			cellHeadStyle.setRightBorderColor(new XSSFColor(new java.awt.Color(221,221,221))); 
			cellHeadStyle.setAlignment(HorizontalAlignment.CENTER);
			cellHeadStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			
			//바디 스타일 저장용 해쉬맵
			HashMap<String, XSSFCellStyle> cellStyles = new HashMap<String, XSSFCellStyle>();
			Sheet[] sheets = new Sheet[sheet_cnt];
			for (int y=0; y < sheet_cnt; y++){//시트 수 만큼 루프
				sheets[y] = wb.createSheet(Integer.toString(y+1));
				for(int rownum = 0; rownum < rowcnt+headcnt; rownum++){//행 만큼 루프
					Row row = sheets[y].createRow(rownum);
					int cellnum = 0;
					
					if(rownum == headcnt && headcnt > 1) {//헤더를 다 그린 후 셀 병합 처리.
						for(int i=0; i<headcnt-1; i++) {
							HashMap<String, Object> row_data = head_excelArray.get(i);
							if(row_data.size()==3) break; //헤더 row의 사이즈가 3보다 작으면 그룹 row가 아니므로 반복문 종료.
							Iterator<?> grid_keys = row_data.keySet().iterator();
							int head_cellnum = 0;
							while(grid_keys.hasNext()) {
								String key = (String)grid_keys.next();
								int len = (Integer)((HashMap<String, Object>) row_data.get(key)).get("LENGTH");
								if(len == 0 ) {
									sheets[y].addMergedRegion(new CellRangeAddress(0,headcnt-1,head_cellnum,head_cellnum));
								}else if(len == 1 ) { // 1그룹에 1개 column인 경우 column을 머지할 필요가 없음.
								}else {
									sheets[y].addMergedRegion(new CellRangeAddress(0,0,head_cellnum,head_cellnum+len-1));
									for(int j=1; j<len; j++){grid_keys.next(); head_cellnum++;}
								}
								head_cellnum++;
							}
						}
					}
					
					if(rownum < headcnt) {//헤더 처리
						HashMap<String, Object> row_data = head_excelArray.get(rownum);
						Iterator<?> grid_keys = row_data.keySet().iterator();
						while(grid_keys.hasNext()) {
							String key = (String)grid_keys.next();
							Cell cell = row.createCell(cellnum);
							cell.setCellValue(String.valueOf((((HashMap<String, Object>)row_data.get(key)).get("TEXT"))));
							sheets[y].setColumnWidth(cellnum, (Integer)((HashMap<String, Object>)row_data.get(key)).get("WIDTH")*35); //엑셀상 적절한 width를 위해 조정
							cell.setCellStyle(cellHeadStyle);
							cellnum++;
						}
					}else {//바디 처리
						
						HashMap<String, Object> row_data = excelArray.get(rownum-headcnt);
						Iterator<?> grid_head_keys = ((HashMap<String, Object>)head_excelArray.get(0)).keySet().iterator();
						String[] grid_body_keys = ((HashMap<String, Object>)excelArray.get(0)).keySet().toArray(new String[excelArray.size()]);
						while(grid_head_keys.hasNext()) {
							//row에 더이상 값이 없을 경우 루프문 종료.
							if(grid_body_keys[cellnum]==null)break;
							
							String body_key = grid_body_keys[cellnum];
							String head_key = (String)grid_head_keys.next();
							String cell_value = "";
							
							//포맷 체크를 위한 선언
							String head_format = (String)((HashMap<String, Object>)((HashMap<String, Object>)head_excelArray.get(0)).get(head_key)).get("FORMAT");
							cell_value = (String)row_data.get(body_key);
							Cell cell = row.createCell(cellnum);
							if(!head_format.equals("") && !cell_value.equals("")) {//편의상 포맷이 있으면 넘버값인 것으로.
								cell.setCellValue(Double.parseDouble(cell_value));
							}else {
								cell.setCellValue(cell_value);
							}
							if(rownum-headcnt == 0) {
								//첫 column 바디 스타일
								XSSFCellStyle cellBodyStyle = (XSSFCellStyle) wb.createCellStyle();
								cellBodyStyle.setBorderTop(BorderStyle.THIN);
								cellBodyStyle.setTopBorderColor(new XSSFColor(new java.awt.Color(221,221,221))); 
								cellBodyStyle.setBorderBottom(BorderStyle.THIN);
								cellBodyStyle.setBottomBorderColor(new XSSFColor(new java.awt.Color(221,221,221))); 
								cellBodyStyle.setBorderLeft(BorderStyle.THIN);
								cellBodyStyle.setLeftBorderColor(new XSSFColor(new java.awt.Color(221,221,221))); 
								cellBodyStyle.setBorderRight(BorderStyle.THIN);
								cellBodyStyle.setRightBorderColor(new XSSFColor(new java.awt.Color(221,221,221))); 
								cellBodyStyle.setVerticalAlignment(VerticalAlignment.CENTER);
								
								//정렬
								String head_align = (String)((HashMap<String, Object>)((HashMap<String, Object>)head_excelArray.get(0)).get(head_key)).get("ALIGN");
								if(head_align.equals("center")) {
									cellBodyStyle.setAlignment(HorizontalAlignment.CENTER);
								}else if(head_align.equals("left")) {
									cellBodyStyle.setAlignment(HorizontalAlignment.LEFT);
								}else if(head_align.equals("right")) {
									cellBodyStyle.setAlignment(HorizontalAlignment.RIGHT);
								}
								//포맷
								if(!head_format.equals("")) {//편의상 포맷이 있으면 넘버값인 것으로.
									XSSFDataFormat format = (XSSFDataFormat)wb.createDataFormat();
									cellBodyStyle.setDataFormat(format.getFormat("#,###0.##0"));
								}
								cellStyles.put(head_key, cellBodyStyle);
							}
							cell.setCellStyle(cellStyles.get(head_key));
							cellnum++;
							if(rownum == rowcnt+headcnt-1) excelArray.get(0).remove(body_key);
						}
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			// 소모시간 반환
			t2 = System.currentTimeMillis();
			System.out.println("spend : " + (t2-t1));
		}
		return wb;
	}
}
