<%
'formbuilder sample - built on bootstrap css
aspl.exec("code/demo_asp/formbuilder/formbuilder.asp")

dim form : set form=new cls_formbuilder
form.listenTo "e","sampleform7"
form.targetDiv="sampleform7"
form.id="hash"

'form-submitted
if form.postback then

	form.doScroll=true
	
	'feedback
	dim md5 : set md5=form.field
	md5.add "type","comment"
	md5.add "html","<strong>MD5 hash:</strong> " & aspL.plugin("md5").md5(aspl.getRequest("yourname"),32)
	md5.add "tag","div"
	md5.add "class","alert alert-dark"	
	
	dim sha256 : set sha256=form.field
	sha256.add "type","comment"
	sha256.add "html","<strong>sha256 hash:</strong> " & aspL.plugin("sha256").sha256(aspl.getRequest("yourname"))
	sha256.add "tag","div"
	sha256.add "class","alert alert-dark"	

end if	

'##########################################################################################

dim yourname : set yourname=form.field
yourname.add "placeholder","password"
yourname.add "type","text"
yourname.add "name","yourname"
yourname.add "class","form-control"
yourname.add "maxlength",50
yourname.add "required",true

'##########################################################################################

dim submit : set submit=form.field
submit.add "style","margin-top:10px"
submit.add "type","submit"
submit.add "name","btnAction"
submit.add "value","Hash"
submit.add "class","btn btn-secondary"

'##########################################################################################

form.build()

%>