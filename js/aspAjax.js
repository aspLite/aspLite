function aspAjax (type,url,data,success) {	
	
	$.ajax({		
		type: type,
		url: url,
		data: data,
		success: success,
		error: aspAjaxError,
	});
};

function aspAjaxError(data) {
	console.log ($(data.responseText).text());
	};