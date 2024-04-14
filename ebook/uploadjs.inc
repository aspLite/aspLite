var i=0
var j=0
var files

$("#upload_aspForm").submit(function(e){

	e.preventDefault()
	
	$("#upload_aspForm").find("input[type=file]").each(function(index, field){
		
		files=field.files
		
		if (files.length>0) {jQueryUpload()}
		
	})
	
});

function jQueryUpload() {	
	
	var file = files[i]	
	
	if (typeof(file) != "undefined") {
		
		var fd = new FormData()			
		fd.append('file', file, file.name)	
		ajax(fd,file)		
		
	}
}

function ajax(fd,file) {

	$.ajax({		
		type: "post",
		method: "post",
		url: aspLiteAjaxHandler + "?asplEvent=uploadfile",
		dataType: "text",
		contentType: false,
		data: fd,
		processData: false, 
		success: function (data) {			
			i++;
			if(i<files.length) {
				jQueryUpload()
				}
			else
				{
					aspAjax('GET','ebook.asp','asplEvent=uploadfb',aspForm)
					}	
				},
		error: function (data) {			
			i++					
			if(i<files.length) {jQueryUpload()}					
			}
		})		

}