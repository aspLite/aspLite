var aspLiteAjaxHandler='default.asp'
var aspLiteAjaxForms = []
var aspLiteFormLooper=0

$(document).ready(function(e) {	

	//bootstrap spinner
	var spinner="<div class='text-center'>"
	spinner+="<div class='spinner-border text-primary spinner-border' role='status'><span class='sr-only'>Loading...</span> </div>"
	spinner+="</div>"
	
	//all divs with class=aspForm will be filled with aspLiteAjaxForms according their id (e-handler)
	$(".aspForm").each(function(){	
		
		//initialize with bootstrap spinners
		$(this).html(spinner)
		
		//I commented the line below out to make the ajax calls sequential
		//this seems to have very little effect on total pageload time
		//but I guess it somehow saves the server processor and it sure
		//limits the number of simultaneous connections to databases (if any)
		
		//aspAjax('GET',aspLiteAjaxHandler,'e=' + $(this).attr('id'),aspForm)			
		
		//console.log($(this).attr('id')) - jQuery reads them top to bottom
		//an array of id's is loaded. it will be used in a second...
		aspLiteAjaxForms.push($(this).attr('id'));
		
	})	
	if (aspLiteAjaxForms.length>0) {
		//let's get this rolling. we start with the first item in the array
		//from there on, the recursive getAspForm will loop through the array and 
		//load all calls sequential
		aspAjax('GET',aspLiteAjaxHandler,'e=' + aspLiteAjaxForms[0],getAspForm)		
	}
})

//recursive execution to avoid server issues
function getAspForm (data) {
	aspForm(data) 
	aspLiteFormLooper++
	//console.log(aspLiteFormLooper)
	if (aspLiteFormLooper < aspLiteAjaxForms.length) {
		//you can even save the server more, by adding some (25) milliseconds of delay before the next call is launched
		setTimeout(function(){ aspAjax('GET',aspLiteAjaxHandler,'e=' + aspLiteAjaxForms[aspLiteFormLooper],getAspForm) }, 25);
		//or if you prefer not to wait...
		//aspAjax('GET',aspLiteAjaxHandler,'e=' + aspLiteAjaxForms[aspLiteFormLooper],getAspForm)
	}
}

$('.ajaxLink').click(function(e) {
		e.preventDefault()
		aspAjax('GET',aspLiteAjaxHandler,'e=' + $(this).attr("data-aspForm"),aspForm)
	})

function aspAjax (type,url,data,success) {	

	$.ajax({		
		type: type,
		url: url,
		dataType: 'json',
		data: data,
		success: success,
		error: aspAjaxError
	});	
};

function aspAjaxError(data) {
	console.log (data);
	}; 
 
 
function aspForm(data) {
			
	//avoid double id's
	$('#' + data.id ).remove()
	
	var aspForm=$('<form>').attr({
		"onsubmit"	: data.onSubmit,
		"style"		: "margin: 0;padding: 0",
		"id"		: data.id,
		"method"	: "post"
		})	
	
	if (data.target == '') {
		
		//form.target is mandatory
		console.log('ERROR: form target is missing!')
		enumerateJson(data)		
	}
		
	//loop through the collection (aspForm) of fieldobjects 

	for(var i = 0; i < data.aspForm.length; i++) {		
	
		var field=data.aspForm[i]
		
		if (typeof field.type == 'undefined') {
			console.log('ERROR: field TYPE is missing!')
			enumerateJson(field)
			continue
		}		
	
		if (field.type=="hidden") {
			$('<input>').attr({
				"type"	: field.type,
				"value"	: field.value,
				"id"	: field.id,				
				"name"	: field.name				
			}).appendTo(aspForm)			
			continue
		}
		
		if (field.type=="link") {
			$('<link>').attr({
				"rel"	: field.rel,
				"href"	: field.href			
			}).appendTo('head')			
			continue
		}

		if (field.type=="button") {
			$('<button>').attr({
				"id"		: field.id,
				"class"		: field.class,
				"style"		: field.style,
				"onclick"	: field.onclick				
			}).html(field.html).appendTo(aspForm)			
			continue
		}			
		
		if (field.type=="style") {
			var style=$('<style>').attr({
				"type"	: "text/css"
			}).html(field.text).appendTo('head')			
			continue
		}
		
		if (field.type=="element") {			
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
	
	//set target (div containing the form)
	$('#' + data.target).html(aspForm)
	
	//scroll to the top of the containing div (offset is used to correct offset from the top of the page, if needed
	if (data.doScroll) {
		$('html,body').animate({scrollTop: $('#' + data.target).offset().top-data.offSet}, 'slow')
	}
	
}

function enumerateJson(p) {
	
	for (var key in p) {
		if (p.hasOwnProperty(key)) {
			console.log(key + " -> " + p[key])
		}
	}
	
}