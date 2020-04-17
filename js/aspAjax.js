var aspAjaxUrl='demo.asp'

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

$(document).ready(function(e) {		
	onPageLoad()
})

$('.aspAjax').click(function(e) {
	e.preventDefault()
	aspAjax('GET',aspAjaxUrl,'text','e=' + this.id,aspAjaxSuccess)	
	scroll()		
})

$('.ajaxForm').submit(function(e) {
	e.preventDefault()
	aspAjax('POST',aspAjaxUrl,'text',$(this).serialize(),aspAjaxSuccess)
	scroll()
})

function onPageLoad () { 
	aspAjax('GET',aspAjaxUrl,'text','e=onload',aspAjaxSuccess)	
}

function aspAjaxSuccess(data) {
	$('#body').html(data)			
}

function scroll() {
	$('html,body').animate({scrollTop: $("#ajaxhandler").offset().top-100}, 'slow')	
}

$('.aspAjaxJSON').click(function(e) {
	e.preventDefault()
	aspAjax('GET',aspAjaxUrl,'json','e=' + this.id,jsonToHTML)
})

function jsonToHTML(data) {

	var records=data.records

	var output='<ul>'

	for(var i = 0; i < records.length; i++) {
	
		var obj = records[i]
		output+="<li>" + obj.sName + ' ('
		output+=obj.iYear + ")</li>";
		
	}
	
	output+='</ul>'	
	
	$('#body').html(output)	
	scroll()
}