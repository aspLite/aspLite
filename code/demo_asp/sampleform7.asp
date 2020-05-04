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
	
	dim hash
	hash=left(aspl.getRequest("haspw"),50)
	
	'feedback
	dim md5 : set md5=form.field
	md5.add "type","comment"
	md5.add "html","<strong>MD5 hash:</strong><textarea rows=""3"" class=""form-control"">" & aspL.plugin("md5").md5(hash,32) & "</textarea>"
	md5.add "tag","div"
	md5.add "class","alert alert-dark"	
	
	dim sha256 : set sha256=form.field
	sha256.add "type","comment"
	sha256.add "html","<strong>sha256 hash:</strong><textarea rows=""3"" class=""form-control"">" & aspL.plugin("sha256").sha256(hash) & "</textarea>"
	sha256.add "tag","div"
	sha256.add "class","alert alert-dark"	

end if	

'##########################################################################################

dim haspw : set haspw=form.field
haspw.add "placeholder","password"
haspw.add "type","text"
haspw.add "name","haspw"
haspw.add "class","form-control"
haspw.add "maxlength",50
haspw.add "required",true

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