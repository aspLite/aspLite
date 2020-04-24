var aspAjaxUrl='demo.asp'

$(document).ready(function(e) {	
	//commented out as this would clash with the default regular handler
	//onPageLoad()
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

$('.aspAjaxJSON2html').click(function(e) {
	e.preventDefault()
	aspAjax('GET',aspAjaxUrl,'json','e=' + this.id,json2HTML)
})

function json2HTML (data) {
    let template = {'<>':'button','style':'margin:10px','class':'btn btn-info','text':'${sName} (${iYear})'};    
    $("#body").html('');
	$("#body").json2html(data.records,template);
	}
    
function jsonToHTML(data) {

	var output='<ul>'

	for(var i = 0; i < data.records.length; i++) {	
		
		output+="<li>" + data.records[i].sName + ' ('
		output+=data.records[i].iYear + ")</li>";
		
	}
	
	output+='</ul>'	
	
	$('#body').html(output)	
	scroll()
}