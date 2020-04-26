<%
on error resume next

if not aspl.isEmp(aspl.getRequest("score")) then

	'this is the place to store the score somehow...

	body="Thanks so much for your <strong>" & aspl.sanitize(aspl.getRequest("score")) & "</strong> stars!"
	
else

	'load the list of stars
	aspL.exec("code/demo_asp/optionList.asp")

	dim starlist : set starlist=new cls_starlist
	'load the template
	body=aspL.loadText("html/demo_asp/rate.resx")
	body=replace(body,"[starlist]",starlist.showSelected("option",""),1,-1,1)
	
end if

aspl.asperror("rating")

aspL.dump body

%>