<%
body=aspl.loadText("default/html/404/default.resx")

'title
body=replace(body,"[title]","Now we're talking!",1,-1,1)

'heading
body=replace(body,"[heading]","Welcome stranger!",1,-1,1)

'body
body=replace(body,"[body]","<p>Welcome to the homepage of this website.</p>",1,-1,1)

aspL.dump(body)

%>