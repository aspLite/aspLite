<%
class cls_asplite_rss

	public template,maxitems,templateBreakPoint,breakpoint,cache
	private xmlDOM,feeditems,child,i,templateItem,item,counter,p_cache
	
	private sub class_initialize
		
		maxitems=6	
		breakpoint=0
		counter=1
		cache=0 'number of seconds this feed will be stored in cache - default 0 (no caching)
		
		template=aspL.loadText(asp_path & "/plugins/rss/template.txt")
		
		templateBreakPoint="<div style=""clear:both""></div>"
	
	end sub

	public function read(url)
	
		if aspL.convertNmbr(cache)<>0 then
		
			read=aspL.getCacheT("asprss" & url,cache)
			
			if not aspL.isEmp(read) then				
				exit function
			end if
		
		end if

		Set xmlDOM = aspL.xmldom(url)
		
		set feeditems=xmlDOM.getElementsByTagName("item")
	
		for i=0 to feeditems.length-1

			templateItem=template
			
			set item=feeditems(i)
			
			for each child in item.childNodes

				select case lcase(child.nodename)
					
					case "title" : templateItem=replace(templateItem,"[TITLE]",child.text,1,-1,1)
						
					case "description" : templateItem=replace(templateItem,"[DESCRIPTION]",child.text,1,-1,1)
						
					case "pubdate" : templateItem=replace(templateItem,"[PUBDATE]",child.text,1,-1,1)
						
					case "link" : templateItem=replace(templateItem,"[LINK]",child.text,1,-1,1)
						
				end select 		
			  
			Next			
			
			'fix missing pubdate
			templateItem=replace(templateItem,"[PUBDATE]","",1,-1,1)
			
			read=read & templateItem
			
			if aspL.convertNmbr(breakpoint)<>0 then				
				if counter=breakpoint then
					read=read & templateBreakPoint
					counter=1
				else
					counter=counter+1
				end if
			end if
			
			if i=maxitems-1 then exit for		
			
		next	

		read=read & templateBreakPoint
		
		'(re)fill cache (aspLite will add the timestamp)
		if aspL.convertNmbr(cache)<>0 then
			aspL.setcache "asprss" & url, read
		else
			'or clear the cache
			aspL.clearcache "asprss" & url
		end if

		Set xmlDOM = Nothing
		Set feeditems = Nothing
		
	end function
	
	public function write
		'tbd
	end function

end class
%>