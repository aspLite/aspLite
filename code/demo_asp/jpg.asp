<%
on error resume next

'get the bootstrap thumbnail template
dim bsTemplate : bsTemplate=aspL.loadText("html/demo_asp/jpg.resx")

dim jpg
set jpg=aspL.plugin("jpg")
jpg.maxsize=500 'px - max: 2560px

'this looks more complex than it is, as this sample is supposed to work in various setups
'by default, this would rather look like jpg.path="/images/img.jpg" where this path is relative to your directory
jpg.path=replace(request.servervariables("path_info"),"demo.asp","",1,-1,1) & aspL_path & "/plugins/jpg/sample.jpg"

dim specialeffects
specialeffects=replace(bsTemplate,"[src]",jpg.src,1,-1,1)
specialeffects=replace(specialeffects,"[caption]","Resizing",1,-1,1)

'color effects
jpg.effect=1 'b/w
	specialeffects=specialeffects & replace(bsTemplate,"[src]",jpg.src,1,-1,1)
	specialeffects=replace(specialeffects,"[caption]","black/white",1,-1,1)
jpg.effect=2 'gray
	specialeffects=specialeffects & replace(bsTemplate,"[src]",jpg.src,1,-1,1)
	specialeffects=replace(specialeffects,"[caption]","gray",1,-1,1)
jpg.effect=3 'sepia
	specialeffects=specialeffects & replace(bsTemplate,"[src]",jpg.src,1,-1,1)
	specialeffects=replace(specialeffects,"[caption]","sepia",1,-1,1)

'crop to rectangle
jpg.effect=0
jpg.square=1 
	specialeffects=specialeffects & replace(bsTemplate,"[src]",jpg.src,1,-1,1)
	specialeffects=replace(specialeffects,"[caption]","square",1,-1,1)

aspL.asperror("jpghandler")

json.dump("<div class=""card-columns"">" & specialeffects & "</div>")

on error goto 0
%>