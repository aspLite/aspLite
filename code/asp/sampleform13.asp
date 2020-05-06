<%
dim jpg
set jpg=aspL.plugin("jpg")
jpg.maxsize=500 'px - max: 2560px
'crop to rectangle
jpg.effect=0
jpg.square=1 

'this looks more complex than it is, as this sample is supposed to work in various setups
'by default, this would rather look like jpg.path="/images/img.jpg" where this path is relative to your directory
jpg.path=replace(request.servervariables("path_info"),"default.asp","",1,-1,1) & aspL_path & "/plugins/jpg/sample.jpg"

dim output
output=replace(aspL.loadText("code/html/jpg.resx"),"[src]",jpg.src,1,-1,1)
output=replace(output,"[caption]","square",1,-1,1)

dim form : set form=aspl.form
form.target="sampleform13"
	
'result
dim feedback : set feedback=form.field("element")
feedback.add "html",output
feedback.add "tag","div"

form.build()

%>