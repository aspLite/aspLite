<%
'load the template file first (including JavScript CDN)
body=aspL.loadText("html/demo_asp/ckeditor5.resx")	

dim mNotes3

mNotes3=aspL.getRequest("mNotes3")

if aspL.isLeeg(mNotes3) then
	mNotes3="<p>Some <strong><u><i>more</i></u></strong> rich text (CKEditor 5).</p>"		
end if

body=replace(body,"[mNotes3]",mNotes3,1,-1,1)

%>