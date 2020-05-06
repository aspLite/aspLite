<%
body=aspl.loadText("code/html/404/default.resx")

'title
body=replace(body,"[title]","Get the code",1,-1,1)

'heading
body=replace(body,"[heading]","Get the code here!",1,-1,1)

'body
body=replace(body,"[body]","<p>Welcome to the get-code page of this website.</p>",1,-1,1)

aspL.dump(body)

%>