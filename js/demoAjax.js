var aspAjaxUrl='demo.asp'

$(document).ready(function(e) {	
	aspAjax('GET',aspAjaxUrl,'e=ajaxhello',aspForm)	
})

$('.aspAjax').click(function(e) {
	e.preventDefault()
	aspAjax('GET',aspAjaxUrl,'e=' + this.id,aspAjaxSuccess)	
	scroll()		
})

$('.aspForm').click(function(e) {
	e.preventDefault()
	aspAjax('GET',aspAjaxUrl,'e=' + this.id,aspForm)		
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

function aspForm(data) {
	
	//avoid double id's
	if (data.id!="") {
		$('#' + data.id ).remove()	
	}	

	var aspForm=$('<form>').attr({
		"onsubmit":"aspAjax('POST','',$(this).serialize(),aspForm);return false;",
		"style":"margin: 0;padding: 0",
		"id":data.id,
		"method":"post"
		})

	for(var i = 0; i < data.aspForm.length; i++) {		
	
		var field=data.aspForm[i]
		
		//console.log(field.type + ' ' + field.name + ' ' + field.value)
	
		if (field.type=="hidden") {
			$('<input>').attr({
				"type"	: field.type,
				"value"	: field.value,
				"id"	: field.id,				
				"name"	: field.name				
			}).appendTo(aspForm)			
			continue
		}
		
		if (field.type=="comment") {			
			$('<' + field.tag + '>').html(field.html).attr({
				"class": field.class,
				"id": field.id,	
				"style": field.style
			}).appendTo(aspForm)			
			continue
		}	
		
		if (field.type=="script") {
			if (field.text!="") {
				$('<script>').text(field.text).appendTo(aspForm)
			}
			else
			{
				$('<script>').attr({
					"src":field.src
				}).appendTo(aspForm)
				
			}
			continue			
		}
	
		if (field.container!="") {
			var formgroup=$('<' + field.container + '>').attr({
				"class": field.containerclass,
				"style": field.containerstyle				
				}).appendTo(aspForm)
		}
		else
		{		
			var formgroup=$('<div>').attr({
				"class": field.containerclass,
				"style": field.containerstyle				
				}).appendTo(aspForm)
		}
		
		if (field.label!=''){
			
			var label=$('<label>').html(field.label).attr({ 
				"for": field.id		
			}).appendTo(formgroup)
				
		}		
		
		if (field.type=="textarea") {			
			$('<textarea>').attr({
				"cols"		: field.cols,
				"rows"		: field.rows,
				"id"		: field.id,	
				"name"		: field.name,
				"class"		: field.class,	
				"style"		: field.style,
				"required"	: field.required					
			}).val(field.value).appendTo(formgroup)		
			continue
		}		
		
		if (field.type=="select") {	
			
			var selectBox=$('<select>').attr({				
				"id"		: field.id,	
				"name"		: field.name,
				"class"		: field.class,
				"style"		: field.style,
				"onchange"	: field.onchange,				
				"required"	: field.required					
			}).val(field.value).appendTo(formgroup)	
	
			//add the options
			var options=field.options
						
			if($.isArray(options)){
			
				//treat as array				
				for(var j = 0; j < options.length; j++) {
					
					if ($.isArray(options[j])){
						//array of arrays
				
						$('<option>').attr({
							
							"value":options[j][0]
							
						}).text(options[j][1]).appendTo(selectBox)
					}
					else
						//array of objects - keyV and pairV expected!
					{
						$('<option>').attr({
						
						"value":options[j].keyV
						
						}).text(options[j].pairV).appendTo(selectBox)		
						
					}
				}				
			}
				
			else
				
			{		
				//treat as JSON object (vbscript dictionary)
				for (var key in options) {						
						
					$('<option>').attr({
						
						"value":key
						
					}).text(options[key]).appendTo(selectBox)						
					
				}				
			}

			//selected value
			selectBox.val(field.value)
			
			continue
		}	
		
		if (field.type=="radio") {	
			
			//add the options
			var options=field.options
			
			var list=$('<ul>').attr({
				
				"style":"list-style:none"
				
			});
			
			for(var j = 0; j < options.length; j++) {	
				
				var item=$('<li>')
				
				var radioB=$('<input>').attr({					
					"type"	: "radio",
					"name"	: field.name,
					"class"	: field.class,
					"style"	: field.style,
					"value"	: options[j][0]					
				}).prop("checked", (field.value==options[j][0])).appendTo(item)
				
				//add label
				$('<span>').html(" " + options[j][1]).appendTo(item)	
				
				item.appendTo(list)
				
			}
			
			list.appendTo(formgroup)
			
			continue
		}		
	
		$('<input>').attr({
			"type"			: field.type,
			"value"			: field.value,			
			"name"			: field.name,
			"class"			: field.class,
			"placeholder"	: field.placeholder,
			"onclick"		: field.onclick,
			"maxlength"		: field.maxlength,
			"id"			: field.id,	
			"style"			: field.style,	
			"required"		: field.required,
			"onchange"		: field.onchange,
			"onclick"		: field.onclick			
		}).appendTo(formgroup)
			
	}	

	//set targetDiv (div containing the form)
	$('#' + data.targetDiv).html(aspForm)
	
	//scroll to the top of the containing div (offset is used to correct offset from the top of the page, if needed
	if (data.doScroll) {
		$('html,body').animate({scrollTop: $('#' + data.targetDiv).offset().top-data.offSet}, 'slow')
	}
	
}