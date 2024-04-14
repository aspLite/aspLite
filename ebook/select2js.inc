$('#js-data-example-ajax').select2({
	placeholder: '', 
	allowClear: true, 	
	ajax: {
    url: 'ebook.asp?asplevent=select2data',
    dataType: 'json',
	delay: 100
  }
});