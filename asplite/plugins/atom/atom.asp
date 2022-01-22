<%
class cls_asplite_atom

	public template, maxitems, cache
	
	private i, xmlDOM, FeedItems, entry, child, updated, published, authorName, authorEmail, authoruri, link, title, id, content, image, summary
	private templateItem, att

	Private Sub Class_Initialize
	
		i=0
		maxitems=10	
		template=aspL.loadText(aspL_path & "/plugins/atom/template.txt")		
		cache=600 'number of seconds this feed will be stored in cache
		
	end sub
	
	public function read(url)
	
		on error resume next		
		
		if aspL.convertNmbr(cache)<>0 then
		
			read=aspL.getCacheT("aspatom" & url,cache)
			
			if not aspL.isEmpty(read) then				
				exit function
			end if
		
		end if		

		Set xmlDOM = aspL.xmldom(url)

		if xmlDOM.getElementsByTagName("feed").length >0 then 
			Set FeedItems = xmlDOM.getElementsByTagName("entry")		
		end if
		
		if xmlDOM.getElementsByTagName("atom:feed").length >0 then 
			Set FeedItems = xmlDOM.getElementsByTagName("atom:entry")		
		end if
		
		for each entry in feeditems		

			templateItem=template
			
			for each child in entry.childnodes
			
				Select case lcase(child.nodeName)
					
					case "updated","atom:updated" : updated = child.text
					
					case "published","atom:published" : published = child.text
						
					case "author","atom:author"					
					
						for each node in child.childnodes
							select case lcase(aspL.convertStr(node.nodeName))
								case "name","atom:name"
									authorName=node.text
								case "email","atom:email"
									authorEmail=node.text
								case "uri","atom:uri"
									authoruri=node.text
							end select
						next						
						
					case "link","atom:link"
						
						if lcase(aspL.convertStr(child.GetAttribute("rel")))="alternate" or aspL.isEmpty(child.GetAttribute("rel")) then
							link = child.GetAttribute("href")
						end if
						
					case "title","atom:title" : title = child.text
						
					case "id","atom:id" : id = child.text
					
					case "content","atom:content" : content = child.text
						
					case "summary","atom:summary" : summary = child.text
						
					case "atom:icon" : image = child.text
						
					case "atom:logo" : image = child.text
						
					case "media:thumbnail"						
					
						for each att in child.attributes
							if lcase(att.name)="url" then
								image = att.text
							end if
						next
						
				End Select		
			
			next
			
			templateItem=replace(templateItem,"[updated]",updated,1,-1,1)
			templateItem=replace(templateItem,"[published]",published,1,-1,1)			
			templateItem=replace(templateItem,"[title]",title,1,-1,1)
			templateItem=replace(templateItem,"[link]",link,1,-1,1)
			templateItem=replace(templateItem,"[summary]",summary,1,-1,1)
			templateItem=replace(templateItem,"[content]",content,1,-1,1)
			templateItem=replace(templateItem,"[image]",image,1,-1,1)
			templateItem=replace(templateItem,"[id]",id,1,-1,1)
			templateItem=replace(templateItem,"[authorname]",authorname,1,-1,1)
			templateItem=replace(templateItem,"[authoremail]",authoremail,1,-1,1)
			templateItem=replace(templateItem,"[authoruri]",authoruri,1,-1,1)
						
			read=read & templateItem
			
			i=i+1
			
			if maxitems=i then exit for
			
		next		
		
		'(re)fill cache (aspLite will add the timestamp)
		if aspL.convertNmbr(cache)<>0 then
			aspL.setcache "aspatom" & url, read
		else
			'or clear the cache
			aspL.clearcache "aspatom" & url
		end if
		
		set xmlDOM=nothing
		set FeedItems=nothing
		
		on error goto 0
	
	end function
	
end class
%>