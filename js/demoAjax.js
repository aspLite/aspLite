var aspAjaxUrl='demo.asp'

$(document).ready(function(e) {	

	var spinner="<div class='text-center'>"
	spinner+="<div class='spinner-border text-secondary spinner-border-sm' role='status'><span class='sr-only'>Loading...</span> </div>"
	spinner+="</div>"
	
	$(".aspForm").each(function(){	
		//initialize with spinners
		$(this).html(spinner)	
		aspAjax('GET',aspAjaxUrl,'e=' + $(this).attr('id'),aspForm)
	})	
})

$('.ajaxLink').click(function(e) {
	e.preventDefault()
	aspAjax('GET',aspAjaxUrl,'e=' + this.id,aspAjaxSuccess)	
	scroll()		
})

$('.ajaxForm').submit(function(e) {	
	e.preventDefault()
	aspAjax('POST',aspAjaxUrl,$(this).serialize(),aspAjaxSuccess)	
	scroll()
})

function aspAjaxSuccess(data) {
	$('#body').html(data.aspl)
}

function scroll() {
	$('html,body').animate({scrollTop: $("#ajax").offset()}, 'slow')	
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
				"style": field.style,
				"onclick": field.onclick				
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