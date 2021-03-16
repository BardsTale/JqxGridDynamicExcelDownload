# JqxGrid Dynamic Excel Download
[![ExportExcel-Jqxgrid](https://img.shields.io/badge/Jqxgrid-ExportExcel-green)](https://github.com/BardsTale/JqxGridDynamicExcelDownload)

JqxGrid의 내장된 Export 기능을 대신하여 DOM 속성과 AbstractXlsxStreamingView를 사용한 스트리밍 방식의<br>
동적 엑셀 다운로드입니다.<br>
대상이 될 JqxGrid의 Column에 따라 동적으로 엑셀 Column을 구성하며 병합된 Column Header도 동적으로 구성합니다.<br>
추가적으로 Column 수를 제한하여 커스터마이징 한 페이징을 하는 경우 페이지 마다 시트를 나눠 동적으로 구성하는 기능도<br>
제공합니다.<br>

SpringBoot 기반의 프로젝트로 예시가 되어있는 코드이며 개발환경이 다를 경우 해당 코드들을 알맞게 수정해주시면 됩니다.<br>
또한 부가적인 UI/UX 옵션을 포함하고 있어 JQuery와 JQuery Cookie 그리고 SweetAlert2를 필요로 하며<br>
핵심 코드는 Vanilla JS로 되어있기 때문에 UI/UX 요소들을 사용하지 않을 경우 해당 라이브러리들을 사용하지 않아도 됩니다.


## 목차

- [필요설치](#필요설치)
    - [프론트엔드](#프론트엔드)
    - [백엔드](#백엔드)
- [사용법](#사용법)
    - [프론트엔드](#프론트엔드_사용법)
    - [백엔드](#백엔드_사용법)
- [관련기술](#관련프로젝트)
- [관리자](#관리자)


## 필요설치

사용을 위해 필요한 설치는 Front-End 방면과 Back-End 방면으로 나뉘어집니다.<br>
※ GitHub에 업로드 된 코드들에 대한 설치는 제외합니다.

### 프론트엔드
이 프로젝트는 JQWidgets의 JqxGrid 기반의 프로젝트이기 때문에 [JQWidgets의 다운로드 및 설치](https://www.jqwidgets.com/download/)가 필요합니다.<br>
NPM을 통한 설치를 원한다면 아래를 참고하세요.<br>
>JQWidgets은 비상업적으론 무료이며 특정 API에 대해선 기능이 제한됩니다.<br>
>설치 시 불필요한 컨텐츠들도 다수 포함하고 있으니 [예시](#예시)를 참고하여 정리하시는 것을 권장드립니다.

```sh
#데모가 포함된 설치
$ npm install jqwidgets-framework

#스크립트만 설치
$ npm install jqwidgets-scripts
```

다음으로 부가적인 UI/UX 구성을 위한 [JQuery](https://www.jqwidgets.com/download/)와 [JQuery Cookie](https://plugins.jquery.com/cookie/)<br>
그리고 [SweetAlert2](https://sweetalert2.github.io/#download)의 설치가 필요합니다.<br>
JQuery와 SweetAlert2는 다음의 CDN과 NPM을 통한 설치를 제공합니다.

1. CDN을 통한 설치
```sh
#JQuery CDN (JQuery Cookie 미포함)
<script src="https://code.jquery.com/jquery-1.12.4.min.js" integrity="sha256-ZosEbRLbNQzLpnKIkEdrPv7lOy9C27hHQ+Xp8a4MxAQ=" crossorigin="anonymous"></script>

#SweetAlert2 CDN
<script src="//cdn.jsdelivr.net/npm/sweetalert2@10"></script>
```

2. NPM을 통한 설치
```sh
#JQuery NPM 설치(JQuery Cookie 미포함)
$ npm install jquery

#SweetAlert2 NPM 설치
$ npm install sweetalert2
```


### 백엔드
AbstractXlsxStreamingView는 Spring 4.2 이후 추가되었으며 Apache POI 3.9 이상 버전과 호환됩니다.<br>
따라서 스트리밍 방식의 SXSSF 하위 클래스를 사용하기 위해 3.9 이상의 [Apache POI](https://poi.apache.org/)가 필요합니다.<br>

1. Gradle을 통한 설치
```sh
dependencies {
...
// https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml
compile group: 'org.apache.poi', name: 'poi-ooxml', version: '4.1.2'
...
}
```

2. Maven을 통한 설치
```sh
...
<!-- https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml -->
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>4.1.2</version>
</dependency>
...
```

## 사용법

사용법은 프론트엔드 방면의 Rest Api 방식의 호출을 통해 가능합니다.

### 프론트엔드_사용법
다양한 방법으로 doExcel(..) 함수를 할당하고 호출하여 사용하면 됩니다.

```sh

//정의된 doExcel 함수.
function doExcel(grid_id,file_name,row_data) {
...
}

//클릭 이벤트에서 doExcel 함수 호출.
function clickEvent(e) 
    //함수의 인자값을 각자 사용하는 방식으로 할당하여 전달.
    /* row_data는 jqxgrid의 getrows 메소드를 사용하여 전달하는 것이 일반적이지만
    * 이 경우 보여지는 column들만 가능하기 때문에 Column 수에 따라 커스터마이징 한 페이징을 하는 경우 알맞지 않습니다.
    */
    doExcel(grid_id, file_name, row_data);
}

// 버튼에 클릭 이벤트 할당
var el = document.getElementById("excelButton");
el.addEventListener("click", clickEvent);
```

### 백엔드_사용법
백엔드 방면에서는 요청을 처리하고 컨트롤러의 ModelAndView 객체를 통해 반환시켜줍니다.

1. 컨트롤러
```sh
@RequestMapping(value="EXCEL", method=RequestMethod.POST)
public ModelAndView postEXCEL(..) {
...
return new ModelAndView(new jqxGridExcelStreamingView(), "model", param_map);
}
```

2. 엑셀을 생성시킬 jqxGridExcelStreamingView 클래스
```sh
@Component
public class JqxGridExcelStreamingView extends AbstractXlsxStreamingView {
    //SXSSFWorkbook 객체 생성
    @Override
    protected SXSSFWorkbook createWorkbook(Map<String, Object> model, HttpServletRequest request) {
        return RenderExcelFromJqxGrid.makeWorkbook(model);
    }
    
    //만들어진 엑셀 파일 처리
    //dispose는 상속받은 AbstractXlsxStreamingView에서 처리된다.
    @Override
	protected void buildExcelDocument(Map<String, Object> model, Workbook workbook, HttpServletRequest request, HttpServletResponse response) throws Exception {
        //응답할 response 객체의 헤더를 설정하여 파일을 다운로드 시킨다.
        ... 코드 생략
    }
}
```

3. 엑셀을 렌더링 할 RenderExcelFromJqxGrid 클래스
```sh
public class RenderExcelFromJqxGrid {
    //실질적인 엑셀 파일을 렌더링하는 코드가 들어간다.
    public static SXSSFWorkbook makeWorkbook(Map<String, Object> model){
        SXSSFWorkbook wb = new SXSSFWorkbook(100); // 100개의 row씩을 메모리에 담고 초과 시 서버의 디스크로 자동 플러쉬한다.
        ...
        return wb;
    }
}
```

## 관련프로젝트

JQWidgets [GitHub](https://github.com/jqwidgets)<br>
JQuery Cookie [GitHub](https://github.com/carhartl/jquery-cookie)


## 관리자

[@BardsTale](https://github.com/BardsTale).
