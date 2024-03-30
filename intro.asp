<!-- #include file="asplite/asplite.asp"-->
<!doctype html>
<html lang="en">
 
 <head>
	
	<title>What exactly is aspLite?</title>
 
    <!-- Required meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
	
	<!-- Bootstrap CSS & JS -->
	<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-QWTKZyjpPEjISv5WaRU9OFeRpok6YctnYmDr5pNlyT2bRjXh0JMhjY6hW+ALEwIH" crossorigin="anonymous">

	<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-YvpcrYf0tY3lHB60NNkmXc5s9fDVZLESaAA55NDzOxhy9GkcIdslK1eN7N6jIeHz" crossorigin="anonymous"></script>

</head>

<body>

<main>
	<div class="container">
	
    <h2>Get started with aspLite</h2>
	
	<p>In its most low-level mode, aspLite is nothing more (or less) than a library of ASP/VBScript classes, functions and subroutines. These can be found in /asplite/asplite.asp.</p>
	
	<p>This page - intro.asp - you're looking at, includes the aspLite framework in it's very first line:<br>
<code>&lt;!-- #include file="asplite/asplite.asp"--&gt;</code></p>

	<div class="alert alert-danger">It's extremly important that this so called include-file is ALWAYS the very first line in your ASP script. I mean: The. Very. First. Line.</div>
	
	<p>Let's have a closer look at <strong>asplite/asplite.asp</strong>. The  first line of that ASP page reads: <strong>&lt;@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%&gt;</strong>. That's probably the most important line in the entire aspLite codebase. What is it doing?
		<ul>
			<li>Set <strong>VBScript</strong> as the default server-side scripting language</li>
			<li>Ensure 100% compatibility with the <strong>UTF-8 character</strong> set, allowing you to use ANY language in your application and avoid very annoying character conversions and/or encoding. Actually, each and every ASP script should start with that one line. Unfortunately, most Classic ASP applications I have seen, did not. With some tragic consequences in most cases.</li>
		</ul> 
		
		</p>
		
	<p>The second line in asplite/asplite.asp reads: <strong>Option Explicit</strong>. This is questionable. I assume you know that by having this line as the very first line in your ASP script, you are forced to declare variables and you're not allowed to declare them more than once. Even though it helps to keep the risks for overwriting certain values under control, it is not a 100% guarantee. Especially when using VBScript's Execute and ExecuteGlobal statements, Option Explicit does not have any effect. And let aspLite use ExecuteGlobal the ... crazy way. So be careful. Make a very good deal with yourself: always declare variables and keep their naming logical and consistent. That's harder than it sounds. Event though in theory you can declare (dim) the same variables (i, j, rs, counter,... are amongst the more popular variable names) in each and every class, function and sub, they DO overwrite eachother. That's no doubt one of the reasons "real" developers never liked VBScript. You never really knew what value (and what type) a VBScript variable held. And if you're not used to that, I can imagine this is driving a developer crazy. It never drove me crazy, as I was pretty crazy already before I wrote my very first line of VBScript code back in 1999.</p>
	
	<div class="alert alert-secondary">I believe the lack of strictly or strongly typed variables in VBScript has caused it's sudden death back in the early 00s and forced the MicroSoft dev-team to come up with the a totally new approach: ASP.NET. It's true that Classic ASP is very prone to a total variable-jungle especially when using lots of include files, not to mention the real disaster scenario when multiple developers had to work on the same codebase. It was nearly impossible to work as a team on a Classic ASP application. However, when working alone - like I always did in my career - it is ... feasible as long as you keep some variable-hygiene into account. PHP was - at that time - lousy typed as well. That did not stop PHP from growing and taking over the web development business as from somewhere mid 2000. My 2 cents.</div>
	
	<p>Let's get on with this.</p>
	
	<p>Right after Option Explicit, 2 ASP files are included.</p>
	<p>The first one: <code>&lt;!-- #include file="config.asp"--&gt;</code>
	</p>
	
	<p>Open that file please.</p>
	
	<p><code>const aspL_path="asplite"</code> lets you decide where exactly you want the aspLite "engine" in your application.</p>
	<p><code>const aspL_debug=true</code> lets you decide whether or not aspLite throws errors. I personally always keep this "true".</p>
		
	<p>The next include file is <code>&lt;!-- #include file="aspForm.asp"--&gt;</code>. That file holds a class, but it is not instantiated. Let's not go into detail right now. Bottom line: it's doing nothing. For now.</p>
			
	<p>The next (and last) thing that asplite/asplite.asp does is create an instance of aspLite (cls_asplite):<br>
	<code>dim aspL<br>set aspL=new cls_asplite</code></p>
	
	<div class="alert alert-info">Do you know that it could also have been:<br><code>dim aspL : set aspL=new cls_asplite</code><br>That would have saved one line of code. Damn.</div>
	
	<p>When creating an instance of a VBScript class, the <strong>Private Sub Class_Initialize()</strong> gets executed (if any). Always. And before anything else.</p>
	<p>Every line in <strong>Class_Initialize</strong> gets executed every time an instance of <strong>cls_asplite</strong> instantiates. So we better think twice before we add an infinite loop over there. Ok, bad joke. But each and every line in <strong>Class_Initialize</strong> is needed and is based on 25 years of experience. </p>
	
	<p>As we're using <strong>Option Explicit</strong>, we're forced - here also - to declare all variable names. In case of a class - rather than "dim" a variable (even though that also "works" in a way), you can use the private or public keyword. Private variables are available within the class only. Public variables (and their values) can be exposed to the outside world. When working alone on apps, you actually don't really need private values. But as aspLite might at some point in time be used by another developer, I guess I can better do it right, and keep most variables in de aspLite class private.<p>
	
	<p>As the code in <strong>Class_Initialize</strong> is always executed when an instance of cls_asplite is created, let's have a close look at what happens, line by line.</p>
	
	<p>
		<code>on error resume next</code><br>This basically tells the ASP-compiler to continue processing the lines below, even in case an error is thrown. But you knew that already, didn't you? What you also HAVE to know, is that this statement needs to be repeated in each and every function or sub. In essence, this "resume next" will be reset at the end of <strong>Class_Initialize</strong>. In the next function, snippet or sub, the ASP compiler will - again - stop executing the script in case an error is thrown. Good to know I guess.</p>
		
		<p>
		<code>startTime					= Timer()</code><br>
		Just because we can and for the fun of it, aspLite holds a timer. startTime will hold the start-time of the script. Let's do this. So far, this page took <strong><%=aspl.printTimer%> milliseconds</strong> to load. That's not much. Having this <code>aspl.printTimer</code> at your fingertips, can help you to isolate badly written code or isolate code that really runs too slow.</p>
		
		<p>
		<code>debug						= aspL_debug</code><br>The value for aspL_debug was set in <code>&lt;!-- #include file="config.asp"--&gt;</code>.</p>	
		
		<p>
		<code>Response.Buffer				= true</code><br>
		Honestly, this is questionable, again. True is the default value. For a reason. If you need to empty the buffer before the ASP script has completely been executed, you can use response.flush and response.clear as you wish. As Response.buffer=true is the default value, this line could have been skipped.		
		</p>
		
		<p>
		<code>Response.CharSet			= "utf-8" 'does not work on IIS5 (Windows 2000 Servers) - comment it out when IIS5 is used!</code><br>	
		Questionable. We already made sure utf-8 is our default charset by adding CODEPAGE="65001" in the first line of asplite/asplite.asp. As the VBScript comments indicate, this line does not work in IIS5 (Windows 2000 Servers). Hence the "On error resume next" above.
		</p>

		<p>
		<code>Response.ContentType		= "text/html"</code><br>
		99% of the output of a web application consists out of text/html. So it's the default contenttype. It gets overwritten if needed though. More about that when discussing the file-serving capabilities of aspLite.
		</p>		
		
		<p>These next four lines ensure that browsers do not cache any output by any ASP page in our project. This is crucial. Back in the late 90s, browser caching was considered useful, as internet connections where slow. Today, you really do not want browsers to cache anything, except the things you really want them to cache (cookies, localStorage, etc).<br>
		<code>Response.CacheControl		= "no-cache"</code><br>
		<code>Response.AddHeader "pragma", "no-cache"</code><br>
		<code>Response.Expires			= -1</code><br>
		<code>Response.ExpiresAbsolute	= Now()-1</code>
		</p>		
		
		<p>		
		<code>Server.ScriptTimeout		= 3600</code><br>
		I agree, this is quite a tolerant value. Our ASP scripts can run for an hour before timing out. Never ever let an ASP page run for an hour. But in some very rare cases, you may have no other option, like when dealing with large file-transfers or migrations.
		</p>
		
		
		<p>The next few lines may sound weird. But they are a crucial part of how aspLite deals with the ASP Request object. In case files are uploaded through a web form, the generic ASP Request collection cannot be used and it even throws an error when called. That's what this little check tries to cover. More about it later. <br>
		<code>'check if a form with enctype="multipart/form-data" was submitted. </code><br>
		<code>'in that case, the request(.form) collection cannot be called as it throws an error</code><br>
		<code>'this is important for getRequest() -> see below</code><br>
		<code>If InStr(Request.ServerVariables("HTTP_CONTENT_TYPE"), "multipart")<>0 Then</code><br>
		<code>	multipart=true</code><br>
		<code>else</code><br>
		<code>	multipart=false</code><br>
		<code>end if</code><br>
		
		
		<p>aspLite comes with an ASP Application-based caching system. Application caching is one of the most underestimated features of Classic ASP. PHP never had a similar function. I have successfully used the ASP Application to store lots and lots (1000's) of values, often in Arrays. Very powerful. Here we set the prefix, so that aspLite will never interfere with yet another caching routine in your (existing) solution.<br>		
		<code>cacheprefix="asplite_"</code><br>
		</p>
		
		<p>aspLite includes some collections (VBScript dictionaries) and instances of other classes. To be able to check whether they are "nothing", they are set to "nothing" to begin with. It's a trick but very effictive.<br> 		
		<code>set plugins			= nothing</code><br>
		<code>set p_fso			= nothing</code><br>
		<code>set p_json			= nothing</code><br>
		<code>set p_randomizer	= nothing</code><br>
		<code>set p_formmessages	= nothing</code><br>
		</p>
		
		<p>
		<code>on error goto 0</code><br>
		This is technically not needed, as we're at the end of the sub anyway. After this, the On Error Resume Next will not have any effect anymore. However, I prefer to clear possible errors before continuing. That's what this line is for. It "whipes" the VBScript Err-object.
		</p>
		
		<p>That was it. This is where we can start using aspLite. </p>
		
		<h2>aspLite as a low-level framework for ASP/VBScript developers</h2>
		
		<p>In its most minimalistic mode, this is it. We now have the timer, the charset, the contenttype, the scripttimeout, the caching and the buffering all set. Good to go. Right?</p>
		
		<p>In a way, yes indeed. These are more than enough reasons to use aspLite for any future ASP/VBScript project.</p>
		
		<p>But come'on. We can do better than this. The very minimal use of aspLite somehow assumes you know about the following aspLite functions. These are some basic aspLite functions that replace some built-in ASP/VBScript functions. In most cases they relate to the major issue the null-value comes with. Nearly all built-in ASP/VBScript functions thrown an error when used with a null-value.</p>
		
		<a name="getRequest"></a><h4>aspl.getRequest(value)</h4>
		<p><code>aspl.getRequest(value)</code> replaces the generic ASP Request object. Unlike the built-in ASP Request Object, it does not throw an error in case a file was uploaded through a form. Pretty much like in the original ASP Request Object, there is a sort order of what exactly is returned: first the form, next the querystring and finally cookies if any.</p>
		
		<p>Example: <a href="intro.asp?q=test#getRequest">Click me</a><br><code>aspl.getRequest</code> returns: <strong><%=aspl.htmlencode(aspl.getRequest("q"))%></strong></p>
		
		<h4>aspl.htmlEncode(value)</h4>
		<p>Previous example made use of another aspLite function: <code>aspl.htmlEncode</code>. Unlike server.htmlEncode(value), aspl.htmlEncode(value) does not throw an error when value is null.</p>
		
		<h4>aspl.isEmpty(value)</h4>
		<p><code>aspl.isEmpty(value)</code> returns true in case value is null, empty or contains blank spaces only. Unlike VBScript's isEmpty(value), aspl.isEmpty(value) does not throw an error when value is null.</p>
		
		<h4>aspl.isNumber(value)</h4>
		<p><code>aspl.isNumber(value)</code> returns true in case value is a number. Unlike VBScript's isNumeric(value), aspl.isNumber(value) does not throw an error when value is null (but returns false in this case).</p>
		
		<h4>aspl.length(value)</h4>
		<p><code>aspl.length(value)</code> returns the length of a value. Unlike VBScript's len(value), aspl.length(value) does not throw an error when value is null (but returns 0 in this case).</p>
		
		<h4>aspl.convertStr(value)</h4>
		<p><code>aspl.convertStr(value)</code> converts value to a string. Unlike VBScript's cstr(value), aspl.convertStr(value) does not throw an error when value is null (but returns an empty string in this case).</p>
		
		<h4>aspl.convertNmbr(value)</h4>
		<p><code>aspl.convertNmbr(value)</code> converts value to a number. In case value is null or a string, 0 is returned.</p>
		
		<h4>aspl.convertBool(value)</h4>
		<p><code>aspl.convertBool(value)</code> returns a boolean (true/false). Unlike VBScript's cbool(value), aspl.convertBool(value) does not throw an error when value is null (but returns false in this case). aspl.convertBool returns true when value is "true", true or 1. In all other cases convertBool returns false.</p>
		
		<h4>aspl.convertNull(value)</h4>
		<p><code>aspl.convertNull(value)</code> converts value to null in case it is 0, empty or null. In case value holds any other number (>0) a number is returned. In case value is a string, a string is returned.</p>
		
		<h4>aspl.convertHtmlDate(date)</h4>
		<p><code>aspl.convertHtmlDate(date)</code> returns a date in the following format: yyyy-mm-dd (needed for the HTML5 date field).</p>
				
		<h4>aspl.pCase(value)</h4>
		<p><code>aspl.pCase(value)</code> converts value to proper case, in addition to VBScript's lCase and uCase functions.</p>
				
		<h4>aspl.curPageURL</h4>
		<p><code>aspl.curPageURL</code> returns the url of the current ASP script, including the protocal, server name and script name.</p>
		
		
		<h4>aspl.fso</h4>
		<p><code>aspl.fso</code> returns an instance of the VBScript FileSystem Object</p>
		
		<h4>aspl.dict</h4>
		<p><code>aspl.dict</code> returns an instance of the VBScript Dictionary Object</p>				
				
				
		<h4>aspl.checkEmailSyntax(strEmail) </h4>
		<p><code>aspl.checkEmailSyntax</code> checks the syntax of an email address - returns true/false</p>
				
		
		
		
		<h4>aspl.log(value)</h4>
		<p>aspLite comes with a basic logger: <code>aspl.log("anything")</code> will write "anything" to aspLite/aspLite.log. The exact time of the logging is included as well. This logging feature is very useful as you can always tell what exactly happens to a variable, or when things go wrong. I often use it to debug certain functions.</p>
		
	</div>
</main>

</body>
</html>