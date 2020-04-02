<%
option explicit

class cls_asp_rss

	public template,maxitems,templateBreakPoint,breakpoint,cache
	private xmlDOM,feeditems,child,i,templateItem,item,counter,p_cache
	
	private sub class_initialize
		
		maxitems=6	
		breakpoint=2
		counter=1
		cache=10 'number of minutes this feed will be stored in cache
		
		template=asp.load(asp_path & "/plugins/rss/template.txt")
		
		templateBreakPoint="<div style=""clear:both""></div>"
	
	end sub

	public function read(url)
	
		if asp.convertGetal(cache)<>0 then

			if isArray(asp.getcache("asprss" & url)) then
			
				'check if it's time to renew
				if DateDiff("n",asp.getcache("asprss" & url)(0),now())<=cache then
					read=asp.getcache ("asprss" & url)(1)					
					if not asp.isLeeg(read) then exit function
				end if
			
			end if
		
		end if

		Set xmlDOM = asp.xmldom(url)
		
		set feeditems=xmlDOM.getElementsByTagName("item")
	
		for i=0 to feeditems.length-1

			templateItem=template
			
			set item=feeditems(i)
			
			for each child in item.childNodes

				select case lcase(child.nodename)
					
					case "title"				
						templateItem=replace(templateItem,"[TITLE]",child.text,1,-1,1)
						
					case "description"
						templateItem=replace(templateItem,"[DESCRIPTION]",child.text,1,-1,1)
						
					case "pubdate"
						templateItem=replace(templateItem,"[PUBDATE]",child.text,1,-1,1)
						
					case "link"
						templateItem=replace(templateItem,"[LINK]",child.text,1,-1,1)
						
				end select 		
			  
			Next			
			
			'fix missing pubdate
			templateItem=replace(templateItem,"[PUBDATE]","",1,-1,1)
			
			read=read & templateItem
			
			if asp.convertGetal(breakpoint)<>0 then				
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
		
		'(re)fill cache
		if asp.convertGetal(cache)<>0 then
			asp.setcache "asprss" & url, array(now(),read)
		else
			asp.clearcache "asprss" & url, array(now(),read)
		end if

		Set xmlDOM = Nothing
		Set feeditems = Nothing
		
	end function
	
	public function write
		'tbd
	end function

end class
%>