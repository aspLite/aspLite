var aspAjaxUrl='demo.asp'

$(document).ready(function(e) {	

	//bootstrap spinner
	var spinner="<div class='text-center'>"
	spinner+="<div class='spinner-border text-secondary spinner-border-sm' role='status'><span class='sr-only'>Loading...</span> </div>"
	spinner+="</div>"
	
	$(".aspForm").each(function(){	
		//initialize with bootstrap spinners
		$(this).html(spinner)	
		aspAjax('GET',aspAjaxUrl,'e=' + $(this).attr('id'),aspForm)
	})	
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
	if (typeof data.id != 'undefined') {
		$('#' + data.id ).remove()	
	}	

	var aspForm=$('<form>').attr({
		"onsubmit"	: data.onSubmit,
		"style"		: "margin: 0;padding: 0",
		"id"		: data.id,
		"method"	: "post"
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
			if (typeof field.text != 'undefined') {
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
	
		if (typeof field.container != 'undefined') {
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
		
		if (typeof field.label != 'undefined') {		
			
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
				"required"	: field.required,
				"multiple"	: field.multiple,
				"size"		: field.size
			}).val(field.value).appendTo(formgroup)	
	
			//add the options
			var options=field.options						
			
			//treat as JSON object (vbscript dictionary)
			for (var key in options) {						
					
				$('<option>').attr({
					
					"value":key
					
				}).text(options[key]).appendTo(selectBox)						
					
			}

			//selected value - array of selected values
			if (typeof field.value != 'undefined') {
				selectBox.val(field.value.split(', '))
			}			
			continue
		}	
		
		if (field.type=="radio") {	
			
			//add the options
			var options=field.options
			
			var list=$('<ul>').attr({				
				"style":"list-style:none"				
			});			
			
			//treat as JSON object (vbscript dictionary)		
			
			for (var key in options) {	
				
				var item=$('<li>')
				
				var radioB=$('<input>').attr({					
					"type"		: "radio",
					"name"		: field.name,
					"class"		: field.class,
					"style"		: field.style,
					"required"	: field.required,
					"value"		: key				
				}).prop("checked", (field.value==key)).appendTo(item)
				
				//add label
				$('<span>').html(" " + options[key]).appendTo(item)	
				
				item.appendTo(list)
				
			}
			
			list.appendTo(formgroup)
			
			continue
		}		
		
		if (field.type=="checkbox") {	
			
			//add the options
			var options=field.options
			
			var list=$('<ul>').attr({				
				"style":"list-style:none"				
			});
			
			//treat as JSON object (vbscript dictionary)
			
			for (var key in options) {	
				
				var item=$('<li>')
				
				var radioB=$('<input>').attr({					
					"type"	: "checkbox",
					"name"	: field.name,
					"class"	: field.class,
					"style"	: field.style,
					"required"	: field.required,
					"value"	: key				
				}).appendTo(item)
				
				if (typeof field.value != 'undefined') {						
					radioB.prop("checked", $.inArray(key.toString(),field.value.split(', '))>=0)
				}				
				
				//add label
				$('<span>').html(" " + options[key]).appendTo(item)	
				
				item.appendTo(list)
				
			}
			
			if (typeof field.value != 'undefined') {
				list.val(field.value.split(', '))
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