function aspAjaxCall (type,url,data,success) {
	$.ajax({
		type: type,
		url: url,
		data: data,
		success: success,
		error: aspAjaxError,
	});
};

function aspAjax(e,obj,onSuccess) {
	
	e.preventDefault();
	
	if (!$(obj).attr('action')=='') { //form is submitted
		var form = $(obj);
		var url = form.attr('action');		
		aspAjaxCall("POST",url,form.serialize(),onSuccess);		
	};
	
	if (!$(obj).attr('href')=='') { //link is clicked
		var link = $(obj);
		var url = link.attr('href');
		aspAjaxCall("GET",url,'',onSuccess);	
	};	
};

function aspAjaxError(data) {
	console.log ($(data.responseText).text());
	};