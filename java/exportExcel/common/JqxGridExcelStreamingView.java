package exportExcel.common;

import java.util.HashMap;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.web.servlet.view.document.AbstractXlsxStreamingView;


@Component
public class JqxGridExcelStreamingView extends AbstractXlsxStreamingView {
	@Override
    protected SXSSFWorkbook createWorkbook(Map<String, Object> model, HttpServletRequest request) {
        return RenderExcelFromJqxGrid.makeWorkbook(model);
    }
	
	@Override
	@SuppressWarnings({"unchecked"})
	protected void buildExcelDocument(Map<String, Object> model, Workbook workbook, HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		HashMap<String, Object> param = (HashMap<String, Object>)model.get("model");
		String filename = param.get("EXLNAM") + ".xlsx"; // 다운로드될 엑셀 파일명 설정.

        String reqCharset = request.getCharacterEncoding(); /*요청의 캐릭터셋 인코딩.*/
        filename = new String(filename.getBytes(reqCharset), "ISO8859-1");
        response.setCharacterEncoding(reqCharset);
        /* 응답 헤더 설정
         * 각종 응답 헤더 설정을 처리한다.
         */
        response.setHeader("Content-Type", "application/octet-stream"); //MIME타입 설정. 텍스트가 아닌 알려지지 않은 파일일 경우.
        response.setHeader("Content-disposition", "attachment;filename=" + filename); //파일명 기입.
        response.setHeader("Set-Cookie", "excelDownCookie=true; path=/"); //다운로드 할 엑셀이 완료됬다는 쿠키 설정. 이를 통해 프로그래스바 관리.
	}
}
