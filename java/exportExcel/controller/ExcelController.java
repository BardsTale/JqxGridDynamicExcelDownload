package exportExcel.controller;

import java.io.IOException;
import java.util.HashMap;

import javax.servlet.http.HttpSession;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import com.ExportExcel.common.JqxGridExcelStreamingView;
import com.fasterxml.jackson.core.JsonGenerationException;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;


@RestController
public class ExcelController {

	@RequestMapping(value="EXCEL", method=RequestMethod.POST)
	public ModelAndView getEXCEL(@RequestParam("param") String params, HttpSession session) throws Exception {
	  //RestController지만 JSON Parser를 특정하지 않았기 때문에 해쉬맵으로 파싱. 특정될 경우 아래의 예외처리를 포함해서 수정.
		ObjectMapper mapper = new ObjectMapper();
		HashMap<String, Object> param_map = new HashMap<String, Object>();
		try {
			param_map = mapper.readValue(params, new TypeReference<HashMap<String, Object>>(){});
		} catch (JsonGenerationException e) {
			e.printStackTrace();
		} catch (JsonMappingException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return new ModelAndView(new jqxGridExcelStreamingView(), "model", param_map);
	}
}
