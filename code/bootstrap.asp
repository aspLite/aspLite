<%
'###########################################################################
'## include these 2 lines to start each codebehind file for security reasons
option explicit  
asp.codebehind() 
'###########################################################################

dim html,titletag,body

'load the template
'resx files are never served to browsers, so they are safer to use
html=asp.load("html/bootstrap.resx")

'hello world plugin example
dim plugin
set plugin=asp.plugin("helloworld")
titletag=plugin.hw()
'or shorter
titletag=asp.plugin("helloworld").hw()
set plugin=nothing

'randomizer plugin example
dim i, randomizer
set randomizer=asp.plugin("randomizer")

body="<p>ASP randomizer-plugin: "
for i=1 to 20
	'generate some random words with random lengths
	body=body & randomizer.randomtext(randomizer.randomnumber(5,10)) & " "
next	
body=body & "</p>"	

'database plugin example
dim db,rs
set db=asp.plugin("accessDB")
db.path="db/sample.mdb"

body=body & "<p>Access database-plugin: " 

set rs=db.execute("select * from person")

while not rs.eof
	body=body & rs("sName") & " "
	rs.movenext
wend 

body=body & "</p>"
%>