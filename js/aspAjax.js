function aspAjax (type,url,dataType,data,success) {	

	$.ajax({		
		type: type,
		url: url,
		dataType: dataType,
		data: data,
		success: success,
		error: aspAjaxError,
	});	
};

function aspAjaxError(data) {
	console.log (data);
	};