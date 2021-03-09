function exportExcelFunc(grid_id, file_name, row_data) {
    let rows = row_data;
    let head_rows = []; 
    if(rows.length == 0){
        warning_swal("엑셀로 출력하실 항목이 없습니다.");
        return;
    }
    let cols = $("#"+grid_id).jqxGrid("columns");
    
    
    wait_swal(); // 화면에 프로그래스 띄우기
    // 1초마다 excelDownCookie 라는 쿠키가 있는지 체크.
    let cookie_check = setInterval(() => {
        if ($.cookie("excelDownCookie") == "true") { 
           $.removeCookie('excelDownCookie', { path: '/' });
           clearInterval(cookie_check);
           Swal.close();
         }
    }, 1000);

    //헤더값 삽입
    let head_Obj = {};
    cols.records.forEach(obj => {
        let head_Property_Obj = {};
        head_Property_Obj['TEXT'] = obj.text;
        head_Property_Obj['WIDTH'] = obj.width;
        head_Property_Obj['ALIGN'] = obj.cellsalign;
        head_Property_Obj['FORMAT'] = obj.cellsformat;
        
        head_Obj[obj.datafield] = head_Property_Obj;
    });
    head_rows.push(head_Obj);
    
    //JqxGrid DOM 프로퍼티에서 그룹값이 있을 경우 기존의 헤더 앞에 삽입
    let has_Parent = cols.records.find(obj => obj.parent != null); // 그룹값이 있는지 체크.
    if(has_Parent !== undefined){
        let parent_Obj = {};
        cols.records.forEach(obj => {
            let parent_Property_Obj = {};
            parent_Property_Obj['TEXT'] = obj.parent==null?obj.text:obj.parent.text;
            parent_Property_Obj['LENGTH'] = obj.parent==null?0:obj.parent.groups.length;
            parent_Property_Obj['WIDTH'] = obj.width;
            parent_Property_Obj['ALIGN'] = obj.cellsalign;
            parent_Property_Obj['FORMAT'] = obj.cellsformat;
            
            parent_Obj[obj.datafield] = parent_Property_Obj;
        });
        head_rows.unshift(parent_Obj);
    }
    
    //엑셀 정보를 담을 data json 생성
    let data     = new Object();
    data.EXLNAM  = file_name;
    data.rows  = rows;
    data.head_rows  = head_rows;
    
    //전송을 위한 가상 form 생성
    var form = document.createElement("form");
    form.method  = "post";
    form.action  = "EXCEL";
    document.body.appendChild(form);
    
    var json = document.createElement("input");
    json.setAttribute("type", "hidden");
    json.setAttribute("name", "param");
    json.setAttribute("contentType", 'application/json');
    json.setAttribute("value", JSON.stringify(data) );
    form.appendChild(json);
    form.submit();
    document.body.removeChild(form);
}