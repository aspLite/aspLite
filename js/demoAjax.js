var aspAjaxUrl='demo.asp'

$(document).ready(function(e) {	
	//commented out as this would clash with the default regular handler
	//onPageLoad()
})

$('.aspAjax').click(function(e) {
	e.preventDefault()
	aspAjax('GET',aspAjaxUrl,'e=' + this.id,aspAjaxSuccess)	
	scroll()		
})

$('.ajaxForm').submit(function(e) {
	e.preventDefault()
	aspAjax('POST',aspAjaxUrl,$(this).serialize(),aspAjaxSuccess)
	scroll()
})

function onPageLoad () { 
	aspAjax('GET',aspAjaxUrl,'e=onload',aspAjaxSuccess)	
}

function aspAjaxSuccess(data) {
	$('#body').html(data.aspl)
}

function scroll() {
	$('html,body').animate({scrollTop: $("#ajaxhandler").offset().top-100}, 'slow')	
}

$('.aspAjaxJSON').click(function(e) {
	e.preventDefault()
	aspAjax('GET',aspAjaxUrl,'e=' + this.id,jsonToHTML)
})

$('.aspAjaxJSON2html').click(function(e) {
	e.preventDefault()
	aspAjax('GET',aspAjaxUrl,'e=' + this.id,json2HTML)
})

function json2HTML (data) {
	
	var template;
	
	$("#body").html('<p><strong>JSON2HTML</strong> - see <a target="_blank" href="http://json2html.com/">http://json2html.com/</a></p>'); 
	$("#body").append('<p>The same <a target="_blank" href="demo.asp?e=json">JSON-data</a> can be visualized in various ways using this library. Fun!</p>'); 
	
	//bootstrap buttons
	$("#body").append('<hr><h5>Buttons</h5>');	
    template = {'<>':'button','onclick':function(){alert("hi!")},'style':'margin:10px','class':'btn btn-info','text':'${sName} (${iYear})'};    
    $("#body").json2html(data.aspl,template);
	
	//html list
	$("#body").append('<hr><h5>HTML lists</h5><ul id="json2html_list"></ul>');	
	template = {'<>':'li','text':'${sName} (${iYear})'};
	$("#json2html_list").json2html(data.aspl,template);
	
	//bootstrap alerts
	$("#body").append('<hr><h5>Bootstrap alerts</h5>');	
	template = {'<>':'div','class':'alert alert-warning','text':'${sName} (${iYear})'};
	$("#body").json2html(data.aspl,template);
	
	//html tables
	$("#body").append('<hr><h5>HTML tables</h5><table class="table table-striped" id="json2html_table"><tbody></tbody></table>');	
	template = {'<>':'tr','html':'<td>${iId}</td><td>${sName}</td><td>${iYear}</td>'};
	$("#json2html_table").json2html(data.aspl,template);
	
	scroll()
	
}
    
function jsonToHTML(data) {

	var output='<ul>'

	for(var i = 0; i < data.aspl.length; i++) {	
		
		output+="<li>" + data.aspl[i].sName + ' ('
		output+=data.aspl[i].iYear + ")</li>";
		
	}
	
	output+='</ul>'	
	
	$('#body').html(output)	
	scroll()
}