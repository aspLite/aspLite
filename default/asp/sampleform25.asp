<%
dim form : set form=aspl.form

if form.postback then

	'button was clicked, we can create the content of the PDF file on the SERVER!
	dim file

	for i=0 to 2500 'PDF with 2500 random strings between 3 and 10 characters long
		file=file & aspl.randomizer.randomtext(aspl.randomizer.randomNumber(3,10)) & " "
	next

	'set javascript line breaks
	dim arr
	arr=split(file," ")
	
	dim line, lines
	lines=0	: file="" : line=""
	
	for i=lbound(arr) to ubound(arr)
		
		line=line & arr(i) & " "
		
		if len(line)>60 then 'lines with max 60 characters
		
			file=file & line & "\n"
			line=""	
			lines=lines + 1
			
			if lines>45 then '45 lines on a page
				file=file & "\addPage"
				lines=0
			end if
			
		end if

	next
	
	file=file & line & "\n"			
	
	dim pages : pages = split(file,"\addPage")
	
	'compile JavaScript 
	dim js
	for i=lbound(pages) to ubound(pages)	
	
		js=js & "doc.text('" & aspl.sanitizeJS(pages(i)) & "', 20, 20);"
		
		if i<ubound(pages) then js=js & "doc.addPage();"
		
	next
	
	dim script : set script = form.field("script")	
	script.add "text","var doc = new jsPDF();doc.setFontSize(14);" & js & "; doc.save('aspLite_" & aspl.randomizer.randomtext(5) & ".pdf')"
	
end if

dim button : set button=form.field("submit")
button.add "value","click to generate pdf"
button.add "class","btn btn-danger"	

form.build()

%>