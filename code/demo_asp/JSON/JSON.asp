<%
'**************************************************************************************************************

'' @CLASSTITLE:		JSON
'' @CREATOR:		Michal Gabrukiewicz (gabru at grafix.at), Michael Rebec
'' @CONTRIBUTORS:	- Cliff Pruitt (opensource at crayoncowboy.com)
''					- Sylvain Lafontaine
''					- Jef Housein
''					- Jeremy Brown
'' @CREATEDON:		2007-04-26 12:46
'' @CDESCRIPTION:	Comes up with functionality for JSON (http://json.org) to use within ASP.
''					Correct escaping of characters, generating JSON Grammer out of ASP datatypes and structures
''					Some examples (all use the <em>toJSON()</em> method but as it is the class' default method it can be left out):
''					<code>
''					<%
''					'simple number
''					output = (new JSON)("myNum", 2, false)
''					'generates {"myNum": 2}
''										
''					'array with different datatypes
''					output = (new JSON)("anArray", array(2, "x", null), true)
''					'generates "anArray": [2, "x", null]
''					'(note: the last parameter was true, thus no surrounding brackets in the result)
''					% >
''					</code>
'' @REQUIRES:		-
'' @OPTIONEXPLICIT:	yes
'' @VERSION:		1.5.1

'**************************************************************************************************************
class JSON

	'private members
	private output, innerCall
	
	'public members
	public toResponse		''[bool] should the generated representation be written directly to the response (using <em>response.write</em>)? default = false
	public recordsetPaging	''[bool] indicates if only the current page should be processed on paged recordsets.
							''e.g. would return only 10 records if <em>RS.pagesize</em> is set to 10. default = false (means that always all records are processed)
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		newGeneration()
		toResponse = false
		recordsetPaging = false
	end sub
	
	'******************************************************************************************
	'' @SDESCRIPTION:	STATIC! takes a given string and makes it JSON valid
	'' @DESCRIPTION:	all characters which needs to be escaped are beeing replaced by their
	''					unicode representation according to the 
	''					RFC4627#2.5 - http://www.ietf.org/rfc/rfc4627.txt?number=4627
	'' @PARAM:			val [string]: value which should be escaped
	'' @RETURN:			[string] JSON valid string
	'******************************************************************************************
	public function escape(val)
		dim cDoubleQuote, cRevSolidus, cSolidus
		cDoubleQuote = &h22
		cRevSolidus = &h5C
		cSolidus = &h2F
		dim i, currentDigit
		for i = 1 to (len(val))
			currentDigit = mid(val, i, 1)
			if ascw(currentDigit) > &h00 and ascw(currentDigit) < &h1F then
				currentDigit = escapequence(currentDigit)
			elseif ascw(currentDigit) >= &hC280 and ascw(currentDigit) <= &hC2BF then
				currentDigit = "\u00" + right(padLeft(hex(ascw(currentDigit) - &hC200), 2, 0), 2)
			elseif ascw(currentDigit) >= &hC380 and ascw(currentDigit) <= &hC3BF then
				currentDigit = "\u00" + right(padLeft(hex(ascw(currentDigit) - &hC2C0), 2, 0), 2)
			else
				select case ascw(currentDigit)
					case cDoubleQuote: currentDigit = escapequence(currentDigit)
					case cRevSolidus: currentDigit = escapequence(currentDigit)
					case cSolidus: currentDigit = escapequence(currentDigit)
				end select
			end if
			escape = escape & currentDigit
		next
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	generates a representation of a name value pair in JSON grammer
	'' @DESCRIPTION:	It generates a name value pair which is represented as <em>{"name": value}</em> in JSON.
	''					the generation is fully recursive. Thus the value can also be a complex datatype (array in dictionary, etc.) e.g.
	''					<code>
	''					<%
	''					set j = new JSON
	''					j.toJSON "n", array(RS, dict, false), false
	''					j.toJSON "n", array(array(), 2, true), false
	''					% >
	''					</code>
	'' @PARAM:			name [string]: name of the value (accessible with javascript afterwards). leave empty to get just the value
	'' @PARAM:			val [variant], [int], [float], [array], [object], [dictionary], [recordset]: value which needs
	''					to be generated. Conversation of the data types is as follows:<br>
	''					- <strong>ASP datatype -> JavaScript datatype</strong>
	''					- NOTHING, NULL -> null
	''					- INT, DOUBLE -> number
	''					- STRING -> string
	''					- BOOLEAN -> bool
	''					- ARRAY -> array
	''					- DICTIONARY -> Represents it as name value pairs. Each key is accessible as property afterwards. json will look like <code>"name": {"key1": "some value", "key2": "other value"}</code>
	''					- <em>multidimensional array</em> -> Generates a 1-dimensional array (flat) with all values of the multidimensional array
	''					- RECORDSET -> array where each row of the recordset represents a field of the array. Each array field has properties according to the column names of the recordset (<strong>lowercase!</strong>) e.g <em>toJSON("r", RS, false)</em> can be accessed afterwards with <em>r[0].id</em>
	''					- <em>request</em> object -> every property and collection (cookies, form, querystring, etc) of the asp request object is exposed as an item of a dictionary. Property names are <strong>lowercase</strong>. e.g. <em>servervariables</em>.
	''					- OBJECT -> name of the type (if unknown type) or all its properties (if class implements <em>reflect()</em> method)
	''					Implement a <strong>reflect()</strong> function if you want your custom classes to be recognized. The function must return
	''					a dictionary where the key holds the property name and the value its value. Example of a reflect function within a User class which has firstname and lastname properties
	''					<code>
	''					<%
	''					function reflect()
	''					.	set reflect = server.createObject("scripting.dictionary")
	''					.	reflect.add "firstname", firstname
	''					.	reflect.add "lastname", lastname
	''					end function
	''					% >
	''					</code>
	''					Example of how to generate a JSON representation of the asp request object and access the <em>HTTP_HOST</em> server variable in JavaScript:
	''					<code>
	''					<script>alert(<%= (new JSON)(empty, request, false) % >.servervariables.HTTP_HOST);</script>
	''					</code>
	'' @PARAM:			nested [bool]: indicates if the name value pair is already nested within another? if yes then the <em>{}</em> are left out.
	'' @RETURN:			[string] returns a JSON representation of the given name value pair
	''					(if toResponse is on then the return is written directly to the response and nothing is returned)
	'******************************************************************************************************************
	public default function toJSON(name, val, nested)
		if not nested and not isEmpty(name) then write("{")
		if not isEmpty(name) then write("""" & escape(name) & """: ")
		generateValue(val)
		if not nested and not isEmpty(name) then write("}")
		toJSON = output
		
		if innerCall = 0 then newGeneration()
	end function
	
	'******************************************************************************************************************
	'* generate 
	'******************************************************************************************************************
	private function generateValue(val)
		if isNull(val) then
			write("null")
		elseif isArray(val) then
			generateArray(val)
		elseif isObject(val) then
			dim tName : tName = typename(val)
			if val is nothing then
				write("null")
			elseif tName = "Dictionary" or tName = "IRequestDictionary" then
				generateDictionary(val)
			elseif tName = "Recordset" then
				generateRecordset(val)
			elseif tName = "IRequest" then
				set req = server.createObject("scripting.dictionary")
				req.add "clientcertificate", val.ClientCertificate
				req.add "cookies", val.cookies
				req.add "form", val.form
				req.add "querystring", val.queryString
				req.add "servervariables", val.serverVariables
				req.add "totalbytes", val.totalBytes
				generateDictionary(req)
			elseif tName = "IStringList" then
				if val.count = 1 then
					toJSON empty, val(1), true
				else
					generateArray(val)
				end if
			else
				generateObject(val)
			end if
		else
			'bool
			dim varTyp
			varTyp = varType(val)
			if varTyp = 11 then
				if val then write("true") else write("false")
			'int, long, byte
			elseif varTyp = 2 or varTyp = 3 or varTyp = 17 or varTyp = 19 then
				write(cLng(val))
			'single, double, currency
			elseif varTyp = 4 or varTyp = 5 or varTyp = 6 or varTyp = 14 then
				write(replace(cDbl(val), ",", "."))
			else
				write("""" & escape(val & "") & """")
			end if
		end if
		generateValue = output
	end function
	
	'******************************************************************************************************************
	'* generateArray 
	'******************************************************************************************************************
	private sub generateArray(val)
		dim item, i
		write("[")
		i = 0
		'the for each allows us to support also multi dimensional arrays
		for each item in val
			if i > 0 then write(",")
			generateValue(item)
			i = i + 1
		next
		write("]")
	end sub
	
	'******************************************************************************************************************
	'* generateDictionary 
	'******************************************************************************************************************
	private sub generateDictionary(val)
		innerCall = innerCall + 1
		if val.count = 0 then
			toJSON empty, null, true
			exit sub
		end if
		dim key, i
		write("{")
		i = 0
		for each key in val
			if i > 0 then write(",")
			toJSON key, val(key), true
			i = i + 1
		next
		write("}")
		innerCall = innerCall - 1
	end sub
	
	'******************************************************************************************************************
	'* generateRecordset 
	'******************************************************************************************************************
	private sub generateRecordset(val)
		dim i, curRow
		write("[")
		curRow = 0
		'recordset.pagesize = -1 means it is not paged.
		while not val.eof and ((recordsetPaging and curRow < val.pageSize) or val.recordCount = -1 or not recordsetPaging)
			innerCall = innerCall + 1
			write("{")
			for i = 0 to val.fields.count - 1
				if i > 0 then write(",")
				toJSON lCase(val.fields(i).name), val.fields(i).value, true
			next
			write("}")
			val.movenext()
			curRow = curRow + 1
			if not val.eof and ((recordsetPaging and curRow < val.pageSize) or val.recordCount = -1 or not recordsetPaging) then write(",")
			innerCall = innerCall - 1
		wend
		write("]")
	end sub
	
	'******************************************************************************************************************
	'* generateObject 
	'******************************************************************************************************************
	private sub generateObject(val)
		dim props
		on error resume next
		set props = val.reflect()
		if err = 0 then
			on error goto 0
			innerCall = innerCall + 1
			toJSON empty, props, true
			innerCall = innerCall - 1
		else
			on error goto 0
			write("""" & escape(typename(val)) & """")
		end if
	end sub
	
	'******************************************************************************************************************
	'* newGeneration 
	'******************************************************************************************************************
	private sub newGeneration()
		output = empty
		innerCall = 0
	end sub
	
	'******************************************************************************************
	'* JsonEscapeSquence 
	'******************************************************************************************
	private function escapequence(digit)
		escapequence = "\u00" + right(padLeft(hex(ascw(digit)), 2, 0), 2)
	end function
	
	'******************************************************************************************
	'* padLeft 
	'******************************************************************************************
	private function padLeft(value, totalLength, paddingChar)
		padLeft = right(clone(paddingChar, totalLength) & value, totalLength)
	end function
	
	'******************************************************************************************
	'* clone 
	'******************************************************************************************
	private function clone(byVal str, n)
		dim i
		for i = 1 to n : clone = clone & str : next
	end function
	
	'******************************************************************************************
	'* write 
	'******************************************************************************************
	private sub write(val)
		if toResponse then
			response.write(val)
		else
			output = output & val
		end if
	end sub

end class
%>