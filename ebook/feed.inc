<%
dim XMLDOM : set XMLDOM=aspL.XMLdom("https://rss.nytimes.com/services/xml/rss/nyt/World.xml")
dim feeditems : set feeditems=XMLDOM.getElementsByTagName("item")
dim i, item, child
dim template, news

for i=0 to feeditems.length-1	

	template="<li>[date]: <a target=""_blank"" href=""[link]"">[text]</a></li>"
	
	set item=feeditems(i)
	
	for each child in item.childNodes

		select case lcase(child.nodename)
			
			case "title" 	: template=replace(template,"[text]",aspl.htmlencode(child.text),1,-1,1)			
				
			case "pubdate" 	: template=replace(template,"[date]",aspl.htmlencode(child.text),1,-1,1)
				
			case "link"  	: template=replace(template,"[link]",aspl.htmlencode(child.text),1,-1,1)
				
		end select 		
	  
	next
	
	news=news & template
	
	if i=5 then exit for
	
next	

response.write "<ul>" & news & "</ul>"

%>