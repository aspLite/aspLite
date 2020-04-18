<%
'you can have multiple ckeditor instances. 

'load the template file first (including JavScript CDN)
body=aspL.loadText("html/demo_asp/ckeditor.resx")	

dim mNotes1,mNotes2
mNotes1=aspL.getRequest("mNotes1")
mNotes2=aspL.getRequest("mNotes2")

if aspL.isLeeg(mNotes1) then
	mNotes1="<p>Some <strong><u><i>rich text</i></u></strong>.</p>"		
end if

if aspL.isLeeg(mNotes2) then
	mNotes2="<p>Some <strong><u><i>more</i></u></strong> rich text.</p>"		
end if

body=replace(body,"[mNotes1]",mNotes1,1,-1,1)
body=replace(body,"[mNotes2]",mNotes2,1,-1,1)
%>