<%
body=aspl.loadText("default/html/404/default.resx")

'title
body=replace(body,"[title]","Contact",1,-1,1)

'heading
body=replace(body,"[heading]","Fill in the contact form:",1,-1,1)

'body
body=replace(body,"[body]","<div class=""aspForm"" id=""sampleform5""></div>",1,-1,1)

aspL.dump(body)

%>