<!-- #include file="aspLite/aspLite.asp"-->
<%
dim asplEvent : asplEvent=aspl.getRequest("asplEvent")

if not aspl.isEmpty(asplEvent) then
	
	'dynamically execute the scriptname in asplEvent
	aspl("ebook/" & asplEvent & ".inc")

end if

function pre(value)
		pre=replace(value,vbtab," ",1,-1,1)
	while instr(pre,"    ")<>0
		pre=replace(pre,"    ","   ",1,-1,1)
	wend 
end function

%>
<!doctype html>
<html lang="en">  
<head>
	
	<title>aspLite: a future for ASP/VBScript</title>
 
    <!-- Required meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
	
	<!-- jQuery  & aspLite JS -->
	<script src="asplite/jquery.js"></script>
	<script>var aspLiteAjaxHandler='ebook.asp'</script>
	<script src="asplite/asplite.js"></script>
	
	<script src="ebook/color-modes.js"></script>
	
	<!-- jQuery UI -->
	<script src="https://code.jquery.com/ui/1.13.2/jquery-ui.min.js" integrity="sha256-lSjKY0/srUM9BE3dPm+c4fBo1dky2v27Gdjm2uoZaL0=" crossorigin="anonymous"></script>
	<link rel="stylesheet" href="https://code.jquery.com/ui/1.13.2/themes/base/jquery-ui.css">
	
	<!-- Bootstrap CSS & JS -->
	<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-QWTKZyjpPEjISv5WaRU9OFeRpok6YctnYmDr5pNlyT2bRjXh0JMhjY6hW+ALEwIH" crossorigin="anonymous">

	<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-YvpcrYf0tY3lHB60NNkmXc5s9fDVZLESaAA55NDzOxhy9GkcIdslK1eN7N6jIeHz" crossorigin="anonymous"></script>
	
	<!-- select2 -->
	<link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
	<script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
		
	<!-- DataTables CSS & JS -->
	<link rel="stylesheet" href="https://cdn.datatables.net/2.0.3/css/dataTables.dataTables.css" />
	<script src="https://cdn.datatables.net/2.0.3/js/dataTables.js"></script>
	
	<!-- jQuery Autocomplete-->	
	<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.devbridge-autocomplete/1.4.11/jquery.autocomplete.min.js"></script>
			
	<style>		
	
		@media print {
		.pagebreak { page-break-before: always; } /* page-break-after works, as well */
		}	
		
		code {font-size:1.05em}
		pre {font-weight:700;padding:25px;}	 
	
		.bd-placeholder-img {
		font-size: 1.125rem;
		text-anchor: middle;
		-webkit-user-select: none;
		-moz-user-select: none;
		user-select: none;
		}

		@media (min-width: 768px) {
		.bd-placeholder-img-lg {
		  font-size: 3.5rem;
		}
		}

		.b-example-divider {
		width: 100%;
		height: 3rem;
		background-color: rgba(0, 0, 0, .1);
		border: solid rgba(0, 0, 0, .15);
		border-width: 1px 0;
		box-shadow: inset 0 .5em 1.5em rgba(0, 0, 0, .1), inset 0 .125em .5em rgba(0, 0, 0, .15);
		}

		.b-example-vr {
		flex-shrink: 0;
		width: 1.5rem;
		height: 100vh;
		}

		.bi {
		vertical-align: -.125em;
		fill: currentColor;
		}

		.nav-scroller {
		position: relative;
		z-index: 2;
		height: 2.75rem;
		overflow-y: hidden;
		}

		.nav-scroller .nav {
		display: flex;
		flex-wrap: nowrap;
		padding-bottom: 1rem;
		margin-top: -1px;
		overflow-x: auto;
		text-align: center;
		white-space: nowrap;
		-webkit-overflow-scrolling: touch;
		}

		.btn-bd-primary {
		--bd-violet-bg: #712cf9;
		--bd-violet-rgb: 112.520718, 44.062154, 249.437846;

		--bs-btn-font-weight: 600;
		--bs-btn-color: var(--bs-white);
		--bs-btn-bg: var(--bd-violet-bg);
		--bs-btn-border-color: var(--bd-violet-bg);
		--bs-btn-hover-color: var(--bs-white);
		--bs-btn-hover-bg: #6528e0;
		--bs-btn-hover-border-color: #6528e0;
		--bs-btn-focus-shadow-rgb: var(--bd-violet-rgb);
		--bs-btn-active-color: var(--bs-btn-hover-color);
		--bs-btn-active-bg: #5a23c8;
		--bs-btn-active-border-color: #5a23c8;
		}

		.bd-mode-toggle {
		z-index: 1500;
		}

		.bd-mode-toggle .dropdown-menu .active .bi {
		display: block !important;
		}
    </style>
	
	<style>
		.autocomplete-suggestions { border: 1px solid #999; background: #FFF; overflow: auto; }
		.autocomplete-suggestion { padding: 2px 5px; color:#000; white-space: nowrap; overflow: hidden; }
		.autocomplete-selected { background: #F0F0F0; }
		.autocomplete-suggestions strong { font-weight: normal; color: #3399FF; }
		.autocomplete-group { padding: 2px 5px; }
		.autocomplete-group strong { display: block; border-bottom: 1px solid #000; }
	</style>
	
	<meta name="twitter:card" content="summary_large_image"/>
	<meta name="twitter:image:src" content="ebook/cover.jpg">
	<meta property="og:image" content="ebook/cover.jpg">
	<meta name="twitter:title" content="aspLite - a future for Classic ASP development">

</head>

<body>
<noscript>
<div class="alert alert-danger no-margin">This application requires JavaScript to be enabled in your browser. JavaScript appears to be disabled in your browser.</div>
</noscript>
<svg xmlns="http://www.w3.org/2000/svg" class="d-none">
<symbol id="check2" viewBox="0 0 16 16">
<path d="M13.854 3.646a.5.5 0 0 1 0 .708l-7 7a.5.5 0 0 1-.708 0l-3.5-3.5a.5.5 0 1 1 .708-.708L6.5 10.293l6.646-6.647a.5.5 0 0 1 .708 0z"/>
</symbol>
<symbol id="circle-half" viewBox="0 0 16 16">
<path d="M8 15A7 7 0 1 0 8 1v14zm0 1A8 8 0 1 1 8 0a8 8 0 0 1 0 16z"/>
</symbol>
<symbol id="moon-stars-fill" viewBox="0 0 16 16">
<path d="M6 .278a.768.768 0 0 1 .08.858 7.208 7.208 0 0 0-.878 3.46c0 4.021 3.278 7.277 7.318 7.277.527 0 1.04-.055 1.533-.16a.787.787 0 0 1 .81.316.733.733 0 0 1-.031.893A8.349 8.349 0 0 1 8.344 16C3.734 16 0 12.286 0 7.71 0 4.266 2.114 1.312 5.124.06A.752.752 0 0 1 6 .278z"/>
<path d="M10.794 3.148a.217.217 0 0 1 .412 0l.387 1.162c.173.518.579.924 1.097 1.097l1.162.387a.217.217 0 0 1 0 .412l-1.162.387a1.734 1.734 0 0 0-1.097 1.097l-.387 1.162a.217.217 0 0 1-.412 0l-.387-1.162A1.734 1.734 0 0 0 9.31 6.593l-1.162-.387a.217.217 0 0 1 0-.412l1.162-.387a1.734 1.734 0 0 0 1.097-1.097l.387-1.162zM13.863.099a.145.145 0 0 1 .274 0l.258.774c.115.346.386.617.732.732l.774.258a.145.145 0 0 1 0 .274l-.774.258a1.156 1.156 0 0 0-.732.732l-.258.774a.145.145 0 0 1-.274 0l-.258-.774a1.156 1.156 0 0 0-.732-.732l-.774-.258a.145.145 0 0 1 0-.274l.774-.258c.346-.115.617-.386.732-.732L13.863.1z"/>
</symbol>
<symbol id="sun-fill" viewBox="0 0 16 16">
<path d="M8 12a4 4 0 1 0 0-8 4 4 0 0 0 0 8zM8 0a.5.5 0 0 1 .5.5v2a.5.5 0 0 1-1 0v-2A.5.5 0 0 1 8 0zm0 13a.5.5 0 0 1 .5.5v2a.5.5 0 0 1-1 0v-2A.5.5 0 0 1 8 13zm8-5a.5.5 0 0 1-.5.5h-2a.5.5 0 0 1 0-1h2a.5.5 0 0 1 .5.5zM3 8a.5.5 0 0 1-.5.5h-2a.5.5 0 0 1 0-1h2A.5.5 0 0 1 3 8zm10.657-5.657a.5.5 0 0 1 0 .707l-1.414 1.415a.5.5 0 1 1-.707-.708l1.414-1.414a.5.5 0 0 1 .707 0zm-9.193 9.193a.5.5 0 0 1 0 .707L3.05 13.657a.5.5 0 0 1-.707-.707l1.414-1.414a.5.5 0 0 1 .707 0zm9.193 2.121a.5.5 0 0 1-.707 0l-1.414-1.414a.5.5 0 0 1 .707-.707l1.414 1.414a.5.5 0 0 1 0 .707zM4.464 4.465a.5.5 0 0 1-.707 0L2.343 3.05a.5.5 0 1 1 .707-.707l1.414 1.414a.5.5 0 0 1 0 .708z"/>
</symbol>
</svg>

<div class="dropdown position-fixed bottom-0 end-0 mb-3 me-3 bd-mode-toggle d-print-none">
<button class="btn btn-bd-primary py-2 dropdown-toggle d-flex align-items-center"
id="bd-theme"
type="button"
aria-expanded="false"
data-bs-toggle="dropdown"
aria-label="Toggle theme (auto)">
<svg class="bi my-1 theme-icon-active" width="1em" height="1em"><use href="#circle-half"></use></svg>
<span class="visually-hidden" id="bd-theme-text">Toggle theme</span>
</button>
<ul class="dropdown-menu dropdown-menu-end shadow" aria-labelledby="bd-theme-text">
<li>
<button type="button" class="dropdown-item d-flex align-items-center" data-bs-theme-value="light" aria-pressed="false">
<svg class="bi me-2 opacity-50" width="1em" height="1em"><use href="#sun-fill"></use></svg>
Light
<svg class="bi ms-auto d-none" width="1em" height="1em"><use href="#check2"></use></svg>
</button>
</li>
<li>
<button type="button" class="dropdown-item d-flex align-items-center" data-bs-theme-value="dark" aria-pressed="false">
<svg class="bi me-2 opacity-50" width="1em" height="1em"><use href="#moon-stars-fill"></use></svg>
Dark
<svg class="bi ms-auto d-none" width="1em" height="1em"><use href="#check2"></use></svg>
</button>
</li>
<li>
<button type="button" class="dropdown-item d-flex align-items-center active" data-bs-theme-value="auto" aria-pressed="true">
<svg class="bi me-2 opacity-50" width="1em" height="1em"><use href="#circle-half"></use></svg>
Auto
<svg class="bi ms-auto d-none" width="1em" height="1em"><use href="#check2"></use></svg>
</button>
</li>
</ul>
</div>

<div class="container-fluid p-5">

<img src="/ebook/cover.jpg" style="width:100%" alt="aspLite: a future for ASP/VBScript">

<div class="mb-4 mt-4 lead">

<p>This book is about Classic ASP and VBScript and how they changed my life in 1998. And how MicroSoft took it back in 2002.</p>

<p>If you want to discuss the various topics covered in this book, do not hesitate to start a discussion on <a href="https://github.com/aspLite" target="_blank">https://github.com/aspLite</a>.</p>

<p class="text-end">Pieter Cooreman, April 2024</p>

</div>

<a name="chapter1"></a>
<h2>Index</h2>

<ol class="list-group list-group-numbered">	
	<li class="list-group-item"><a class="link" href="#chapter2">Preface</a></li>
	<li class="list-group-item"><a class="link" href="#chapter3">Story behind aspLite</a></li>
	<li class="list-group-item"><a class="link" href="#chapter4">Where does aspLite fit in?</a></li>
	<li class="list-group-item"><a class="link" href="#chapter5">How aspLite lead to asplForm</a></li>
	<li class="list-group-item"><a class="link" href="#chapter6">aspLite plug-ins</a></li>
	<li class="list-group-item"><a class="link" href="#chapter7">Excursions</a></li>
	<li class="list-group-item"><a class="link" href="#chapter8">Final notes</a></li>
</ol>

<div class="pagebreak"></div>
<a name="chapter2"></a>
<h2>Preface</h2>

<p>Back in 2002, many thousands - difficult to say how many exactly - of ASP/VBScript developers were left alone in the woods by a very small team of MS developers who chose to build ASP.NET and throw ASP and VBScript in the bin. ASP got canceled, from one day to the other. That's how it felt back then, and that's what really happened looking back on it today.</p>

<p>I sometimes try to picture it like this: In 2002, Microsoft had just had a great couple of years. Windows 95, 98, 98SE, 2000 Professional, NT Server, 2000 Advanced Server. Not to mention Windows Me and the various XP editions. MicroSoft also had Internet Explorer, Office, SQL Server and many other products. These were all great products that were loved by basically any gamer, student, school, non-for-profit and  all businesses. Between 1997 and 2002, MicroSoft was the leading software company worldwide, without any doubt. They sure could afford to make mistakes and make some wrong decisions. And so they did. That's how Classic ASP/VBScript ended up in the bin, without much ado. Just like that. But - as always - pride comes before a fall.</p>

<div class="mt-4 mb-4 p-4 text-bg-dark">

<p>Ever since 2002, Classic ASP/VBScript developers have lived in the Dark Middle Ages. We're being ignored, laughed at, made fun of and expelled from MicroSoft's Promised Land.NET. We're loose ends, lone wolves. This hurts. This has to stop. Rather sooner than later.</p>

</div>

<p>ASP (Active Server Pages) can be combined with multiple Active Scripting languages in IIS. VBScript has always been the most popular choice (by far). JScript and Perlscript were the two others, but they were barely used. 99% of all ASP web applications used and still use VBScript as their preferred Active Scripting language.</p>

<div class="alert alert-success">As a 10 year old kid, I was a lucky owner of a <strong>Commodore 64</strong>. C64's had the programming language <strong>Basic</strong> on board. I remember how I spent evenings trying to code my way around very "basic" tasks and animations. VBScript did have a lot in common with that Basic on my Commodore 64. When I started my developer career in 1998, using VBScript felt like coming home after all these years.</div>

<p>But I was not the only one. ASP/VBScript were very popular web development and scripting technologies between 1997 and 2002. They were MicroSoft's first answer to another rapidly growing web technology and competitor: PHP and Apache servers. We all know how that turned out. However, most businesses preferred ASP over PHP back then. ASP was running faster, integrated well with other MicroSoft products (Active Directory, SQL Server, Office, ...), had a rapidly growing community and supported (multi-threaded) COM-extensions. But above all, ASP/VBScript was very easy to learn, even for people like me, without a degree in computer science whatsoever. VBScript was a basic, visual, case insensitive and lousy typed scripting language. And I loved it.</p>

<p>Even though ASP was a huge success, MicroSoft developers decided to pull the plug and amuse themselves with... how many ... well what the heck, 25 versions of ASP.NET. ASP.NET never was nowhere near a popular web development technology. ASP/VBScript was. In its early years, ASP.NET was left far behind by PHP frameworks. Today full stack JavaScript frameworks have taken over.</p>

<p>PHP developers had more luck. Between 2000 and 2010 (about the same period of the rise and fall of ASP.NET and Windows Servers) a handful of PHP developer frameworks gained popularity amongst web developers (Zend, CakePHP, Symfony, Laravel, ...). They were there to stay. Everybody felt that. Unlike ASP.NET, these PHP frameworks were <i>built</i> with PHP, they did not want to <i>replace</i> it. And PHP ran on less expensive hosts and servers.</p>

<p>ASP/VBScript developers also needed a developer framework back in 2002. They did not need ASP.NET. A developer framework for ASP/VBScript should have taken care of various shortcomings in ASP/VBScript (better support for dealing with files - pdf, jpg, zip, etc), workarounds for known bugs or issues, facilitate code behind, become event-driven, bring a solution for the spaghetti-coding, improve coding habits, url-rewriting (MVC!), increase scalability and security. Last but not least, ASP/VBScript needed a framework in order not to reinvent the wheel each time a new project was to be developed. But that framework never happened.</p>

<p>Modern and popular web development frameworks like Django and Vue are also facilitating the so-called spaghetti-way of coding. It does not prevent those frameworks to gain a large following amongst both beginning and more experienced developers. Same was 100% true for Classic ASP back then.</p>

<p>By the time ASP.NET was stable (not before 2.0 in 2005), most talented developers had given up on .NET and went developing plug-ins for WordPress, Joomla or Drupal, or they developed one or the other social network. Using PHP. For me it was very simple: By then, I had 100s of customers who were using my Classic ASP applications. I built a hosting business around ASP/VBScript somewhere between 2000 and 2005. I did not want to refactor 10000s of lines of code just because MicroSoft wanted me to. So I got stuck with Classic ASP. And even today in 2024, I'm still hosting 100s of web applications solely relying on ASP/VBScript.</p>

<p><strong>Classic ASP was considered sloppy, buggy and weak. A technology for lesser Gods.</strong></p>

<p>At least that's what MicroSoft developers wanted us to believe back then. And I sure identify as a lesser developer God. But the Classic ASP community did have some individuals who did a great job with some very powerful scripts. I remember a pure ASP/VBScript upload class by Lewis Moten that I used a lot in <a class="link" target="_blank" href="https://quickersite.com">QuickerSite</a> and other projects. It got fine-tuned and Unicode-proof by other developers over the years. Others developed very useful encryption classes for Classic ASP (Sha256.asp by Phil Fresle in 2001). I also remember a very useful Captcha image generator-class by Emir Tüzül in 2006. Later on we also saw a few popular JSON-classes.</p>

<p>But the Classic ASP community also saw some great attempts to develop true frameworks. Michal Gabrukiewicz did a great job with <a class="link" target="_blank" href="http://www.asp-AJAXed.org/">asp-AJAXed</a> but sadly passed away in 2009. More about Michal later. He played an important role when developing aspLite, even though he passed away 10 years before I started developing it. <a href="https://www.codeproject.com/Articles/7922/Classic-ASP-Framework-2-0-Make-your-Classic-ASP-co" target="_blank">CLASP</a> by Christian Caldereon was another great .NET-alike framework for Classic ASP. Others came up with Classic ASP MVC implementations, each one of them relying on the very useful 404-rewriting-trick. All in all, developer frameworks for ASP/VBScript developers did not last long and were often developed (and used) by a single developer. Apparently it was very hard to build a community around Classic ASP back then. I still blame MicroSoft. MicroSoft has given its own ASP/VBScript tandem very bad press back in those days. Nobody dared to stick his neck out by building a durable and community-driven ASP/VBScript framework back in 2003-10. I still regret that. We only needed 1 widely used Classic ASP (MVC) framework back in 2000. And we could and should have built it. But MicroSoft developers did not want us to have fun with that, they wanted to have fun themselves. So they created .NET.</p>

<p><strong>No painless nor hassle-free upgrade-path to ASP.NET</strong></p>

<p>Back in those days, ASP/VBScript developers did not have a painless nor hassle-free upgrade path to ASP.NET. VBScript was not supported by ASP.NET. Oh yes, for small web applications, one or two ASP pages, it was doable. MicroSoft even provided some automated conversion tools. But they were very limited. In 2003, by the time ASP.NET 1.1 fixed some very annoying bugs in 1.0, I was dealing with 3 extremely large ASP/VBScript code bases. Tons of includes files, classes, functions and routines. It was impossible to refactor them to ASP.NET without spending at least a year. And for what reason? ASP.NET did not offer much extra compared to ASP/VBScript at that time. In 2002, ASP.NET was often considered <i>too little too late</i> by seasoned MicroSoft developers. But what the heck, I didn't have the time for all that. I was building a business around my developing, selling and communication skills. And I really didn't need ASP.NET for that.</p>

<p>There was (and still is!) another very good reason to stay away from .NET back in 2002-2005! .NET web applications always needed much more RAM compared to Classic ASP apps. I mean: MUCH more RAM. An average Classic ASP application (still today) only needs a maximum of 30MB of RAM - that is - when done the aspLite-way (or MVC), using the one-entry-point method. Therefore it was perfectly possible to host between 50 and 100 Classic ASP applications on a single Windows Server with 4GB of RAM (a very common RAM-amount for production servers back in 2000). ASP.NET web applications easily need 10 times more. For a business owner this is a no-brainer. If you are/were into ASP.NET development, you need more and stronger hardware. The cost of building a SAAS hosting-business - like I have been doing all my life - is at least two/three times higher when using ASP.NET.</p>

<p>As soon as a technology gets introduced and gains a large user base - as was the case for ASP/VBScript between 1997 and 2002 - it's impossible to stop it. People will always be prepared to live with its limitations, work their way around them or learn to live with them. That's exactly what I did back then, and I'm still doing now with aspLite. And I'm not alone. It's a human thing. And we're all humans after all. MicroSoft misjudged that.</p>

<p>We're 25 years later now. There are still loads (millions) of Classic ASP/VBScript applications out there, most of them serving dynamic websites for more than 25 years now. Still MicroSoft refuses to clearly communicate about the EOL policy of Classic ASP. There IS no EOL policy for ASP/VBScript. So after all these years, we - ASP/VBScript developers - are STILL left in the woods. Alone. Without even the slightest clue on when exactly MicroSoft will pull the plug.</p>

<p>I learned to live with that even though I still feel sad about it. In 2020 I decided to develop a new framework for ASP/VBScript developers: <a class="link" href="https://aspLite.com" target="_blank">aspLite</a>. I will be working on aspLite for the rest of my life. I somehow love this technology. And it just won't end. Classic ASP is fun! And fun is key.</p>

<p><strong>Some kudos for MicroSoft though</strong></p>

<p>Sure, it's not all bad news either. MicroSoft still offers Classic ASP through ISAPI in IIS and it made sure that all our ASP applications just kept on running smoothly like they did back in 2000. They even run better now, thanks to today's faster hardware. MicroSoft did not mess up Classic ASP. They did not add things to it, but they also did not screw things up. That's a plus.</p>

<p>We also saw some interesting attempts to provide an IDE for Classic ASP developers. We had two flavors of "Webmatrix", an all-in-one IDE for both MicroSoft and Open Source technologies. I liked both versions of Webmatrix a lot. But they both suddenly disappeared at some point in time.</p>

<p>So far, all Visual Studio editions supported intellisense and code-completion for Classic ASP. And even all IIS Express-versions support Classic ASP/VBScript to the bone. Today, even the Windows 10 and 11 Home editions come with a <i>full version of IIS</i>. In a way, nowadays it's much easier to start developing in ASP/VBScript than it was back then. You needed Windows 2000 Professional or a Server back in 2000. Today you only need to know how to <i>enable</i> ASP to get it up and running. Unfortunately, very few Windows users still know how to do that, and they couldn't care less. This illustrates how things turned out for MicroSoft. In 2000, companies spent quite some money on Windows 2000 Pro licenses. Today, MicroSoft ships its entire development framework and its dependencies for free. But even that never turned the tide.</p>

<p>Another pleasant evolution is that - over the years - MicroSoft embraced Open Source technology. At some point in time, it was easier to install both WordPress, Joomla and Drupal on any Windows host than it was to get ASP applications up and running. The wonderful Web Platform Installer however, was retired in 2022 - again - for inexplicable reasons. Probably because it didn't bring any money in and only pushed developers even more towards Open Source developer frameworks. I'm curious about how long <strong>Visual Studio Code</strong> will survive. The most popular plug-ins for VSC are about ... Python. Not quite a MicroSoft technology either.</p>

<div class="pagebreak"></div>
<a name="chapter3"></a>
<h2>Story behind aspLite</h2>

<p>The story behind aspLite tells the story of my career and the way I dived into web development about 24 years ago.</p><p><strong>The early days</strong></p><p>24 years ago - in 2000 - I was 28 years old. I was young and eager to learn to code. I had no degree in computer science whatsoever, but I picked up a lot from colleagues very quickly. Developing web applications quickly became an obsession. I developed all sorts of applications &nbsp;- both as a professional web developer, and for a hobby. Back in those days, I happened to work for a company that specialized in developing COM components for e-commerce websites. I was not part of the RnD team actually building these components. I was part of the support-team, implementing them for customers. We used Classic ASP and VBScript. What else did you think?</p><p><strong>Visual Basic Scripting - VBScript&nbsp;</strong></p>

<p>Rather than use the COM components made by that RnD team, I quickly realized that it was actually much easier (and much quicker) to develop custom classes in VBScript to fully meet the customer's requirements. Actually, using those COM components slowed down our applications and the development cycles. So in the end, we didn't use them. It somehow meant the end of that company. And I was in it for something... well, a lot actually. When I left that company in 2002, I took its biggest customer with me and started my own company. Shame on me. But hey... that's life. I was a 30 year old entrepreneur after all. I simply had to do it.</p>

<p>VBScript was the first programming language I learned to use. Full stack developers will claim that VBScript is useless and cannot be called a serious programming language. I disagree. VBScript is <strong>visual </strong>(easy to read/write, case-insensitive coding, no nested curly braces {{{{{}}}}} - I mean...), <strong>basic</strong> (easy to understand, no complex statements) and <strong>scripted</strong> (just-in-time, no compilation). These 3 properties made a <em><strong>huge success</strong></em> of VBScript back in 1997-2002. VBScript can also be used together with ActiveX Data Objects (ADO) - a high-level, easy-to-use interface to OLE databases (Access, SQL Server, Oracle, etc). ADO is what made VBScript a success. And it still does.</p>

<p><strong>QuickerSite</strong></p>

<p>After having coded my way through basically any type of web application somewhere between 2000 and 2007, I decided to come up with a CMS (Content Management System) that combined all my best scripts and coding habits that had passed the test of time so far: <a href="https://www.quickersite.com/" target="_blank">QuickerSite</a>. QuickerSite was a success, especially in its early years. Only one year after its initial release in 2007, it was translated into 11 languages, including Danish, Hebrew, Italian, Turkish and Swedish. In 2008 QuickerSite was used by about 1000 users worldwide, who created at around 6000 QuickerSites in total. It was as if a lot of ASP coders all over the world had been waiting for a CMS developed in pure ASP/VBScript.&nbsp;</p>

<p>Between 2007 and 2014 I built a hosting business around QuickerSite. At its peak, I hosted 1200 QuickerSites on a single dedicated Dell server with very basic specs (3GB RAM, a slow 120 GB SATA disk and a single CPU). But it worked. It rocked. And all this time, I was a one-man-band. Nobody else but me, myself and I were dealing with everything related to my business: selling, developing, designing, hosting, invoicing, mailing. And everything else related to QuickerSite. So yes, I sure was a bit of a lone wolf back then. And as much as Microsoft forced me to switch over to ASP.NET, I did the exact opposite and stayed with Classic ASP. That's me.</p>

<p>Developing, selling, hosting and supporting QuickerSite for a wide variety of customers (and hosting conditions) thaugt me even more about ASP/VBScript, its caveats AND its strengths. Developing QuickerSite was no doubt the best time of my life. I still host many QuickerSites today on a Windows 2019 Server in the cloud. And I made some life-time friends.</p>

<p><strong>The 2010 paradigm shift</strong></p>

<p>By the time QuickerSite had grown mature - somewhere in 2010, there were quite a few things going on: HTML5, CSS3 and JavaScript frameworks got adopted by the WWW, mobile devices (phones and tablets) were rapidly gaining in popularity, social media took over our lives, and last but not least - by then - a large group of developers adopted open source solutions and frameworks developed in PHP/MySQL. Nowadays, JavaScript/CSS frameworks like Bootstrap (Twitter), React (Facebook), Angular (Google) and Node.js are completely dominating the web development business. While Microsoft was trying hard to keep up with these ever changing circumstances - by releasing dozens of ASP.NET versions and editions - Classic ASP developers were slowly and silently being ignored and left alone in the woods. Many of them are retired by now, or they are no longer actively developing (new) Classic ASP applications. I regret that a lot, because all this time - and still today - ASP/VBScript has been a perfect fit for these JavaScript frameworks, for instance by providing very straight-forward database access (ADO), dealing with binary files (uploading/streaming) or by simply delivering very useful AJAX, XML and JSON integrations. All that is where aspLite is about.</p>

<div class="mt-4 mb-4 p-4  text-bg-primary">
<p><strong>Trip down memory lane</strong></p>

<p>When Jesse James Garret "invented" AJAX in 2005 - and Google made a success of it - some Classic ASP/VBScript developers must have thought by themselves: what the heck, we did AJAX back in 1998 already! We used something named <strong>Remote Scripting</strong>. It was AJAX avant-la-lettre. We loaded a Java Applet (it took about 3 seconds to load) and next we were able to provide AJAX functionality in pure Classic ASP/VBScript. Back in 1998. This technique worked fine on both Internet Explorer and Netscape, the only two browsers that mattered back then.</p>
</div>

<p><strong>PHP and ASP.NET libraries in Classic ASP</strong></p>

<p>As Classic ASP is a dead-end street anyway, it may be a good idea to do some neighborhood shopping. Why not use some PHP or ASP.NET libraries in Classic ASP? I've been doing that for many years already. I use .NET's web.config files to configure url rewriting (http-&gt;https), set custom error handling (404 catch) and set default documents (default.asp). I also developed a single VB.NET page that takes care of (unobtrusive) server-side image-resizing/cropping and recoloring. I added that aspx-page to the jpg-plugin in aspLite (jpeg.aspx). The idea of using resx extensions for HTML files is another (ab)use of .NET. All in all it's not much, but it is something.</p>

<p>In the past I have also successfully implemented PHP's built-in zipper (ZipArchive()) and a plugin named mPDF to create zip -and pdf files on the server from within a Classic ASP application. Worked like a charm. As these libraries reside on the server, they can safely be implemented with internal ServerHTTPRequests from a Classic ASP environment.</p>

<p><strong>Concerns about the future of Classic ASP</strong></p>

<p>Some developers may ask themselves: "<em>For how long are we going to be able to host Classic ASP applications on Windows Servers</em>"? Windows Server 2022 fully supports all Classic ASP/VBScript components. Windows Server 2025 will not remove support for ASP/VBScript either.</p>

<p>As for the end-of-life policy of ASP, MicroSoft has published <a href="https://learn.microsoft.com/en-us/troubleshoot/developer/webapps/iis/general/asp-support-windows" target="_blank">this very short article</a> on the 25th of January 2022. In short: both ASP and IIS support lifetimes are tied to the support lifecycle of the host operating system, in most cases a Windows Server. The support lifecycle for Windows 2022 is Oct 14 2031. The support lifecycle for Windows 2025 will probably be 3 years later, somewhere the end of 2034. In all honesty, I believe we will be able to host Classic ASP/VBScript web applications for at least 10 more years in optimal conditions (somewhere until 2035). After that, things might become tricky. </p>

<div class="mt-4 mb-4 p-4 text-bg-warning">
<p><strong>VBScript is about to be removed from Windows OS. Really?</strong></p>
<p>In Oct 2023, MicroSoft announced that <a href="https://learn.microsoft.com/en-us/windows/whats-new/deprecated-features-resources" class="link" target="_blank" style="color:#000">VBScript will retire</a>. MicroSoft says: <i>"VBScript will be available as a feature on demand before being retired in future Windows releases. Initially, the VBScript feature on demand will be pre installed to allow for uninterrupted use while you prepare for the retirement of VBScript."</i> </p>

<p>This announcement is - again - very confusing. Will this also affect Classic ASP/VBScript web applications relying on server-side VBScript in IIS? Or will VBScript only be removed as a client-based scripting language in Windows Scripting Host? These are two completely different things. MicroSoft's communication is not very clear on this matter.</p>
</div>

<p><strong>Suggestion: turn ASP/VBScript into Open Source</strong></p>

<p>The best way forward for Classic ASP is to turn it into an open source project, pretty much like Microsoft did with <a href="https://github.com/dotnet/aspnetcore" target="_blank">ASP.NET Core</a>. Given today's success of scripted languages like JavaScript, Ruby and PHP, VBScript still has the potential to reach a different kind of developer. Some developers (like me) will NEVER be able to read through dozens of nested curly braces. It's just visually too poor for me. But that does not mean we are unable to write poetry in ASP/VBScript. I'll keep on trying anyway.</p>

<div class="mt-4 mb-4 p-4  text-bg-primary">

<h4>What and who I'm doing this for...</h4>

<p>I often receive messages like the ones below. It's striking how many of us Classic ASP developers share the same feeling.&nbsp;Left alone in the woods with our ASP-scripts and functions that passed the test of time (some of them are working flawlessly for 25 years now, on all Windows Servers).</p>

<p><i>"I am moved by your back stories with the development of aspLite. I have been an ASP web developer since 2004 and still using asp Classic as for my freelance projects. I can relate on the struggle on having shared hosting upto know and I would love to learn more about AWS EC2 instance and own a hosting server like what you did. I'm happy to know that there are still believing in the power of Classic ASP."</i></p>

<p><i>"I like your work! Contrary to what the muppets say, Classic ASP is not dead. There is nothing that one cannot do with ASP. In fact I have written desktop applications in ASP... cannot do that with PHP. You may recall SOOP? It had a huge following and a lot of devs for a lot of useful plug-ins. Not sure what happened with that even though I was one of their plugin devs. Perhaps it was the cost of web hosting vs Apache servers and WordPress. Good luck with your CMS."</i></p>

<p><i>"I've just found your web site for the first time today and I just wanted to say thank you for doing it. I've been using Classic ASP since the late 90s and still use it today for projects even though I've been jeered and sneered at for doing so. It's great to see there are other people still out there using it for real and still flying the flag :0)"</i></p>

<p><i>"I found your page by chance, because I was looking for something in ASP that would work with the "DataTables | Table plugin for jQuery"... congratulations for maintaining this site. Hope you are still working with ASP...thanks."</i></p>

<p>Out of all the emails I have ever received, I will always remember this one. It's about QuickerSite, but it also applies to Classic ASP/VBScript... in a way:</p>

<p><i>"Pieter, many years ago I came across Quickersite and I paid I think for something (I hope I was a payer and not just a downloader) and I was overwhelmed with how amazing it all was. I thought I'd found the holy grail for me being a provider of web services... but my history has followed a similar trajectory to QS and I'm now wandering the plains with a lot of other buffaloes wondering what to do next. I read your posts occasionally as they are, and for a weird reason - it gives me hope. Hope because what you had going here was SO ON THE MARK. But the world rotated a little in another direction and twitter/FB / Instablah and all the other mindless platforms took the collective consciousness away from individual effort and relevant opinion, such that QS no longer had a home. ... Why? Who knows, the world is just different now and windows of opportunity open and close for the most inexplicable of reasons. This could of course be used as an equal answer to: "Why does humanity exist"? A: "For inexplicable reasons". So; no more Concorde, no more digital watches, no more moon landings and no more QS. It is in good company and I loved it."</i></p>

</div>

<h4>Some final notes before diving into the code</h4>

<p><strong>IDE</strong></p>

<p>Classic ASP developers have been lacking a dedicated IDE (Integrated Development Environment) for some time now. I therefore prefer to use <a href="https://notepad-plus-plus.org/" target="_blank">Notepad++</a>. It's free, lightning fast, reliable and it's very easy to use. Notepad++ also comes with some basic code-completion and intellisense for ASP. Notepad++ can also be freely installed and used on any Windows Server. I often use it to search for specific texts and strings in over 1 million files. No problem at all. Others may want to use <a href="https://code.visualstudio.com/" target="_blank">Visual Studio Code</a> instead.</p>

<p><strong>Visual Studio Code</strong> is MicroSoft's Open Source and free code-editor. Its community is offering some very useful extensions for Classic ASP/VBScript developers. They usually include syntax highlighting, intellisense, and code navigation for VBScript inside Active Server Pages (ASP) files. When combined with one of the available IIS Express Extensions, Visual Studio Code can be considered a useful IDE (Integrated development environment) for Classic ASP development.</p>

<p><strong>Where to host ASP sites today?</strong></p>

<p>A final note on hosting. I have personally never used or liked <em>shared </em>ASP hosting solutions of any type. From day one, I use my own server. In 2004 I bought my own. In 2010 I migrated to the cloud. Today (2024), I use an&nbsp;<a href="https://aws.amazon.com/ec2/" target="_blank">AWS EC2 instance</a>. Very satisfying so far. Windows Servers got much cheaper over the years. You get a lot of "Windows Server" for somewhere 100$/month these days, including backups, data transfer and lots of disk space. In my opinion, if you're into ASP development, you're better off managing a Windows Server <em>yourself</em>, rather than rely on shared hosting. There are very few people left who can assist you with ASP hosting issues. We're on our own. But that shouldn't be a problem as Windows Servers are very easy to set up, deploy and maintain. As an ASP developer you need to know the basics of setting up backups, firewall, IIS, mail server and security on Windows Servers after all. On top of that, when you dive into ASP development, you may have to install specific COM software or you may want to prefer a specific setup to facilitate code-reuse. This requires full access to a server. Both Microsoft Azure, Google Cloud and Amazon Web Services offer free-tier solutions that allow you to test-drive a basic setup for a year. It's really worthwhile looking into these solutions.</p>

<p><strong>Should I learn Classic ASP/VBScript in 2024?</strong></p>

<p>Hell no.</p>

<p>If you're young and you're willing or able to dive into new technologies, do not learn Classic ASP or VBScript. If you want to become a web developer (by extension a mobile app developer) - today or in the future - learn Python and/or JavaScript. However, if you are already using Classic ASP (and plan to keep on doing so), it's always a good idea to learn new things. Even after all these years, there's always something new to learn about ASP. aspLite is living proof of that.</p>

<p>That said, learning ASP actually means: learn VBScript! There is not much to learn about Classic ASP after all, except for 6 objects with each a handful of properties and methods: Application, Session, Response, Request, Error and Server.  VBScript however is the preferred (and default) programming language for most ASP developers. Even though VBScript is no longer being developed by Microsoft, it is still used a lot, not only by web developers, but also to automate tasks on Windows servers.</p>

<p>Out of many online resources you find when searching for "learn Classic ASP", I personally liked '<a href="https://asplite.com/assets/files/A%20Practical%20Guide%20to%20ASP%203.0.pdf" target="_blank">A Practical Guide to Microsoft Active Server Pages 3.0</a>' by Manas Tungare a lot. He was so kind to allow me to redistribute this guide on this website. These 70 well written pages are about everything you need to get started using Classic ASP and VBScript.</p>

Furthermore, W3Schools.com (that website is developed in Classic ASP) is probably the best online resource to start as a Classic ASP/VBScript web developer. Before you learn Classic ASP, you must learn HTML and CSS (and grab some basic JavaScript as well). Make sure to also check their very complete VBScript and ADO reference. ADO can be used to access databases from Classic ASP pages. </p>

<p><strong>About this e-book</strong></p>

<p>This book ain't really a book. It's an ASP script, using aspLite as its preferred framework. It appears that the easiest way to write a book about aspLite is using it while I'm trying to explain it. All the code that I include and comment, is also executed whenever there is a "live preview". This book uses an Access database, so you would have to enable 32bit applications for your IIS application pool.</p>

<div class="pagebreak"></div>
<a name="chapter4"></a>
<h2>Where does aspLite fit in?</h2>

<div class="mt-4 mb-4 p-4  text-bg-success">

<p>24 years ago, Classic ASP was about Request() and Response(). Ins and outs. Simply put.</p>
<p>Today, any server-side web framework is still about just that.</p>
<p>For everything else you can - in 2024 - rely on JavaScript frameworks. No more need for spaghetti-ASP or HTML server-controls.</p>

</div>

<p>aspLite is a lightweight framework for ASP/VBScript developers. It can help to develop better ASP/VBScript applications. aspLite does not reinvent, replace or rewrite ASP components. It only tries to come up with a new way of using some of them.</p>

<p>ASP/VBScript (also known as "Classic ASP") has been available for more than 25 years now and I believe it still makes sense to share this approach. After all, Classic ASP has proven to be reliable, fast and secure. </p>

<p>But that's not all. Today's web applications use JavaScript frameworks (Angular, Vue, React, jQuery, etc). Often only data (JSON) is transmitted back and forth to the server. That's why a server-side technology better stays small, fast and simple. Hence the name "aspLite".</p>

<p>An example. The <a href="https://demo.asplite.com" target="_blank">aspLite demo</a> site ships with a (fully functional) Classic ASP implementation of DataTables. This wonderful (free-to-use) widget has all it takes to offer ... datatables, including client-side sorting, searching and paging. There is only very little ASP code involved. Quite fascinating!</p>

<p>The vast majority of these JavaScript plug-ins are free to use, backed by 100's of contributors and they work in all (current) browsers, on all devices. How amazing is that?! Ever since the adoption of HTML5, CSS3 and ECMAScript 5 (somewhere between 2009 and 2012), client-side JS/CSS/HTML frameworks have become very popular. Much more popular than their server-sided predecessors (ASP(.NET), PHP, etc). Today, the most starred repositories on GitHub are all about (learning) JavaScript.</p>

<p>The <a href="https://demo.asplite.com" target="_blank">aspLite demo</a> puts the following front-end HTML/CSS/JS frameworks/plug-ins at work: Bootstrap, jQuery, jQuery UI Datepicker, JSZip, SheetJS, jsPDF, CkEditor, CodeMirror and DataTables. </p>

<div class="alert alert-success"><p>aspLite does not rely on 3rd party COM components. It only relies on the built-in VBScript components. Therefore aspLite works on basically each and every shared hosting solution out there, but also on each and every Windows host with ASP installed.</p></div>

<h4>Install aspLite as a virtual directory for code-reuse reasons?</h4>

<p>If you're lucky enough to manage a Windows Server yourself, you may want to install aspLite as a <i>virtual directory</i> in your IIS website. This way you can use one aspLite codebase on an unlimited number of websites on your server. If you use a network location, you can even share one aspLite codebase amongst multiple servers.</p>

<p>When setup this way, the first line in your ASP page would read:</p>
<pre class="alert alert-light">&lt;!-- #include virtual="aspLite/aspLite.asp"--&gt;</pre>
<p>(in case you name the virtual directory "aspLite")</p>


<h4>Getting started</h4>

<p>In its most low-level mode, aspLite is nothing more (or less) than a library of ASP/VBScript classes, functions and subroutines. They can be found in /aspLite/aspLite.asp. I will go through all of them later on in this article.</p>

<p>This page - <strong>ebook.asp</strong> - you're looking at, includes the aspLite framework in it's very first line:</p>

<pre class="alert alert-light">&lt;!-- #include file="aspLite/aspLite.asp"--&gt;</pre>

<div class="alert alert-danger">It's extremely important that this so-called include-file is ALWAYS the very first line in your ASP script.</div>

<p>Let's have a closer look at <strong>aspLite/aspLite.asp</strong>. The  first line of that ASP page reads:</p>

<pre class="alert alert-light">&lt;@LANGUAGE="VBScript" CODEPAGE="65001"%&gt;</pre>

<p>That's probably the most important line in the entire aspLite code base. What is it doing?</p>

<ul>
	<li>Set <strong>VBScript</strong> as the default server-side scripting language</li>
	<li>Ensure 100% compatibility with the <strong>UTF-8 character</strong> (Unicode) set, allowing you to use ANY language in your application and avoid very annoying character conversions and/or buggy encoding. Actually, each and every ASP script should start with that one line. Unfortunately, most Classic ASP applications I have seen, did not. With some tragic consequences in most cases.</li>
</ul>

<p>The second line in aspLite/aspLite.asp reads:</p>

<pre class="alert alert-light">Option Explicit</pre>

<p>I assume you know that by having this line as the very first line in your ASP script, you are forced to declare variables and you're not allowed to declare them more than once. Even though it helps to keep the risks of overwriting certain values under control, it is not a 100% guarantee. Especially when using VBScript's Execute and ExecuteGlobal statements, Option Explicit does not have any effect. So be careful. Make a very good deal with yourself: always declare variables and keep their naming logical and consistent. That's harder than it sounds. Even though in theory you can declare (dim) the same variables (i, j, rs, counter,... are amongst the more popular variable names) in each and every class, function and sub, they DO overwrite each other. That's no doubt one of the reasons "real" developers never liked VBScript. You never really knew what value (and what type) a VBScript variable held. And if you're not used to that, I can imagine this is driving a developer crazy.</p>

<div class="mt-4 mb-4 p-4  text-bg-dark">
I believe that the lack of strictly or strongly typed variables in VBScript has caused its sudden death back in the early 00s and forced the MicroSoft dev-team to come up with their totally new approach: ASP.NET. It's true that Classic ASP is very prone to a total variable-jungle especially when using lots of include files, not to mention the real disaster scenario when multiple developers had to work on the same code base. It was nearly impossible to work as a team on a Classic ASP application. However, when working alone - like I always did in my career - it is ... feasible as long as you keep some variable-hygiene into account. PHP was - at that time - lousy typed as well. That did not stop PHP from growing and taking over the web development business as from somewhere mid 2000. My 2 cents.
</div>

<p>Right after Option Explicit, 2 ASP files are included.</p>
<p>The first one:</p>

<pre class="alert alert-light">&lt;!-- #include file="config.asp"--&gt;</pre>

<p><code>const aspL_path="aspLite"</code> lets you decide where exactly you want the aspLite "engine" in your application.</p>
<p><code>const aspL_debug=true</code> lets you decide whether or not aspLite throws errors. I personally always keep this "true".</p>

<p>The next include file is:</p>

<pre class="alert alert-light">&lt;!-- #include file="aspForm.asp"--&gt;</pre>

<p>That file holds a class, but it is not instantiated. Let's not go into detail right now. Bottom line: it's doing nothing. For now.</p>
	
<p>The next (and last) thing that aspLite/aspLite.asp does is create an instance of aspLite (cls_aspLite):</p>
<pre class="alert alert-light">dim aspL<br>set aspL=new cls_aspLite</pre>

<div class="alert alert-info">Do you know that it could also have been:</p>

<pre class="alert alert-light">dim aspL : set aspL=new cls_aspLite</pre>

<p>In ASP/VBScript, you can separate commands with a colon. That would have saved one line of code.</p></div>

<p>When creating an instance of a VBScript class, the <strong>Private Sub Class_Initialize()</strong> gets executed (if any). Always. And before anything else.</p>
<p>Every line in <strong>Class_Initialize</strong> gets executed every time an instance of <strong>cls_aspLite</strong> instantiates. So we better think twice before we add an infinite loop over there. Ok, bad joke. But each and every line in <strong>Class_Initialize</strong> is thoughtfully written.</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/classinit.inc")))%></pre>

<p>As we're using <strong>Option Explicit</strong>, we're forced - here also - to declare all variable names. In case of a class - rather than "dim" a variable (even though that also "works" in a way), you can use the private or public keyword. Private variables are available within the class only. Public variables (and their values) can be exposed to the outside world. When working alone on apps, you actually don't really need private values. But as aspLite might at some point in time be used by another developer, I guess I can better do it right, and keep most variables in the aspLite class private.<p>

<p>As the code in <strong>Class_Initialize</strong> is always executed when an instance of cls_aspLite is created, let's have a close look at what happens, line by line.</p>

<pre class="alert alert-light">on error resume next</pre>

<p>This basically tells the ASP-compiler to continue processing the lines below, even in case an error is thrown. But you knew that already, didn't you? What you also have to know is that this statement needs to be repeated in each and every function or sub. In essence, this "resume next" will be reset at the end of <strong>Class_Initialize</strong>. In the next function, snippet or sub, the ASP compiler will - again - stop executing the script in case an error is thrown. Good to know I guess.</p>

<pre class="alert alert-light">startTime = Timer()</pre>

<p>Just because we can and for the fun of it, aspLite holds a timer. startTime will hold the start-time of the script. Let's do this. So far, this page took <strong><%=aspl.printTimer%> milliseconds</strong> to load. That's not much. Having this <code>aspl.printTimer</code> at your fingertips, can help you to isolate badly written code or isolate code that really runs too slow.</p>

<pre class="alert alert-light">debug = aspL_debug</pre>

<p>The value for aspL_debug was set in <code>&lt;!-- #include file="config.asp"--&gt;</code>.</p>	

<pre class="alert alert-light">Response.Buffer = true</pre>

<p>Honestly, this is questionable, again. True is the default value anyway. For a reason. If you need to empty the buffer before the ASP script has completely been executed, you can use response.flush and response.clear as often as you wish. As Response.buffer=true is the default value, this line could have been skipped.</p>

<pre class="alert alert-light">Response.CharSet = "utf-8" 'does not work on IIS5 (Windows 2000 Servers) - comment it out when IIS5 is used!</pre>

<p>Questionable. We already made sure utf-8 is our default charset by adding CODEPAGE="65001" in the first line of aspLite/aspLite.asp. As the VBScript comments indicate, this line does not work in IIS5 (Windows 2000 Servers). Hence the "On error resume next" above.</p>

<pre class="alert alert-light">Response.ContentType	= "text/html"</pre>

<p>99% of the output of a web application consists of text/html. So it's the default content type. It can easily be overwritten if needed though. More about that when discussing the file-serving capabilities of aspLite.</p>		

<p>These next four lines ensure that browsers do not cache any output by any ASP page in our project. This is crucial. Back in the late 90s, browser caching was considered useful, as internet connections were slow. Today, you really do not want browsers to cache anything, except the things you really want them to cache (cookies, localStorage, etc).</p>

<pre class="alert alert-light">
Response.CacheControl = "no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires = -1
Response.ExpiresAbsolute = Now()-1
</pre>		

<pre class="alert alert-light">Server.ScriptTimeout	= 3600</pre>

<p>I agree, this is quite a tolerant value. Our ASP scripts can run for an hour before timing out. Never ever let an ASP page run for an hour. But in some very rare cases, you may have no other option, like when dealing with large file-transfers or occasional migrations or synchronizations.</p>

<p>The next few lines may sound weird. But they are a crucial part of how aspLite deals with the ASP Request object. In case files are uploaded through a web form, the generic ASP Request collection cannot be used and it even throws an error when called. That's what this little check tries to cover. More about it later.</p>


<pre class="alert alert-light">
'check if a form with enctype="multipart/form-data" was submitted. 
'in that case, the request(.form) collection cannot be called as it throws an error
'this is important for getRequest() -> see below
If InStr(Request.ServerVariables("HTTP_CONTENT_TYPE"), "multipart")<>0 Then
	multipart=true
else
	multipart=false
end if
</pre>

<p>aspLite comes with an ASP Application-based caching system. Application caching is one of the most underestimated features of Classic ASP. PHP never had a similar function. I have successfully used the ASP Application to store lots and lots (1000's) of values, often in Arrays. Very powerful. Here we set the prefix, so that aspLite will never interfere with yet another caching routine in your (existing) solution.</p>

<pre class="alert alert-light">
cacheprefix="aspLite_"
</pre>

<p>aspLite includes some collections (VBScript dictionaries) and instances of other classes. To be able to check whether they are "nothing", they are set to "nothing" to begin with. It's a trick but very effective.</p>

<pre class="alert alert-light">
set plugins	= nothing
set p_fso = nothing
set p_JSON = nothing
set p_randomizer = nothing
set p_formmessages = nothing
</pre>

<pre class="alert alert-light">on error goto 0</pre>

<p>This is technically not needed, as we're at the end of the sub anyway. After this, the On Error Resume Next will not have any effect anymore. However, I prefer to clear possible errors before continuing. That's what this line is for. It "wipes" the VBScript Err-object.</p>

<p>That was it. This is where we can start using aspLite. </p>

<h3>Common methods in aspLite</h3>

<p>In its most minimalistic mode, this is it. We now have the timer, the charset, the content type, the script timeout, the caching and the buffering all set. Good to go. Right?</p>

<p>In a way, yes indeed. These are more than enough reasons to use aspLite for any future ASP/VBScript project.</p>

<p>However, there is much more to aspLite than just some settings. The very minimal use of aspLite somehow assumes you know about the following aspLite methods. These are some basic aspLite functions that replace or extend some built-in ASP/VBScript functions. In most cases they relate to the major issue the null-value comes with. Nearly all built-in ASP/VBScript functions throw an error when used with null.</p>

<h4>aspl.exec(filepath) and aspl.executeASP(text)</h4>

<p>Executes the content of a Unicode file (relative filepath) or in case of aspl.executeASP - plain text - as ASP/VBScript. After the ASP code has been loaded, it is available in the namespace of the ASP script.</p>

<p>aspLite heavily relies and facilitates VBScript's <strong>executeGlobal</strong> function. In short: executeGlobal allows to "import" ASP code in an ASP page's namespace. The "imported" code does not even have to reside on the same server. It can be located anywhere on the internet. ExecuteGlobal imports any text and treats it as ASP code.</p>

<p>We all know the idea behind include-files, but executeGlobal somehow includes ASP code the "extreme" way.</p>

<p><strong>An example:</strong></p>

<pre class="alert alert-light">
&lt;!-- #include file="aspLite/aspLite.asp"--&gt;
&lt;%
dim hw : hw="https://demo.aspLite.com/ebook/helloworld.txt"
aspl.executeASP(aspl.XMLhttp(hw,false))
%&gt;
</pre>

<p>The file <a href="https://demo.aspLite.com/ebook/helloworld.txt" target="_blank">https://demo.aspLite.com/default/html/helloworld.txt</a> does actually exist and serves some valid ASP code. aspLite loads the file and executes it. I'm not sure MicroSoft ever intended to use executeGlobal this way. But it works. I'm sure you can see lots of great opportunities, but also lots of extremely dangerous scenarios.</p>

<p><strong>Live example:</strong></p>

<pre class="alert alert-light"><%aspl.executeASP(aspl.XMLhttp("https://demo.aspLite.com/ebook/helloworld.txt",false))%></pre>

<p>Another example. Imagine you have this default.asp:</p>

<pre class="alert alert-light">
&lt;!-- #include file="aspLite/aspLite.asp"--&gt;
&lt;%
aspl.exec("scripts/" &amp; aspl.getRequest("script") &amp; ".inc")
%&gt;
</pre>

<p>The folder "scripts" holds 3 files:</p>
<ul>
<li>1.inc - <code>response.write "hello1"</code></li>
<li>2.inc - <code>response.write "hello2"</code></li>
<li>3.inc - <code>response.write "hello3"</code></li>
</ul>

<p>Finally, browse to:</p>

<ul>
<li><a href="http://localhost/?script=1">http://localhost/?script=1</a></li>
<li><a href="http://localhost/?script=2">http://localhost/?script=2</a></li>
<li><a href="http://localhost/?script=3">http://localhost/?script=3</a></li>
</ul>

<p>The one and only default.asp runs 3 different "imported" scripts. Unlike when using include-directives, the imported ASP code is dynamically loaded (or not). You decide which ASP script has to load and which does not. This keeps RAM usage under control. But above all, it facilitates an amazingly well structured ASP code base. For each event you can have your dedicated classes, functions, subs and program flow. And still you would only be using only one ASP page (default.asp) as your main entry-point. That one ASP page is able to serve each and every module of your ASP application. This begins to smell like MVC.</p>

<div class="mt-4 mb-4 p-4  text-bg-success">
<p><strong>CDN for Classic ASP/VBScript classes?</strong></p>
<p>The Hello-World script above (https://demo.aspLite.com/ebook/helloworld.txt) could have led - back in 2000 - to another idea. But it didn't. Anyway, the idea would be to set up a CDN (Content Delivery Network) serving well-written ASP-classes. Pretty much like JavaScript frameworks rely on CDN, Classic ASP/VBScript CDN could have been set up in a very similar way. It would not be the browser loading a CDN file, but the server loading full-blown (compressed) classes to deal with (less) common tasks and functions. THAT's what we needed in 2002. We did not need ASP.NET. Also, this CDN would have been a major step towards Open Source. Still today, it would be a great help for Classic ASP developers. We don't even need a hosting solution. All we need is a couple of megabytes of cloud storage. This idea of CDN for Classic ASP goes beyond the scope of this book. It's something worth experimenting with however.</p>
</div>

<div class="mt-4 mb-4 p-4  text-bg-warning">
<p><strong>Be careful with aspl.exec() and aspl.executeASP().</strong></p>
<p>Dynamically importing ASP code with aspl.exec or aspl.executeASP is very powerful. It can be used to create plug-ins, import remote code, keep your codebase strictly structured, etc. Unfortunately there is no way to trace errors in ASP code that is imported this way. Using regular include files will return appropriate error messages like we're used to and we would always need to during development. The underlying VBScript function <code>executeGlobal</code> is to blame. Make sure to use <code>aspl.exec</code> and <code>aspl.executeASP</code> with care and only for small ASP snippets or well-tested ASP code.</p>
</div>

<a name="getRequest"></a><h4>aspl.getRequest(value)</h4>

<p><code>aspl.getRequest(value)</code> replaces the generic ASP Request object. Unlike the built-in ASP Request Object, it does not throw an error in case a file was uploaded through a form. Pretty much like in the original ASP Request Object, there is a sort order of what exactly is returned: first the form, next the querystring and finally cookies if any.</p>

<h4>aspl.htmlEncode(value)</h4>

<p>Unlike server.htmlEncode(value), <code>aspl.htmlEncode</code> does not throw an error when value is null.</p>

<div class="alert alert-danger">It can't be said enough: ALWAYS htmlencode ANY input from users. I mean: also first names, last names and street numbers. <strong>Anything</strong>. It's the first and most efficient protection against XSS (Cross-Site-Scripting).</div>

<h4>aspl.sqli(value)</h4>
<p>Protects your SQL queries against SQLinjection. In essence, single quotes in value are replaced with two single quotes.</p>

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
<p>Example:</p>

<pre class="alert alert-light">&lt;input type="date" value=&lt;%=aspl.convertHtmlDate(date)%&gt; /&gt;</pre> 

<p>returns:</p> 

<p><input type="date" value=<%=aspl.convertHtmlDate(date)%> /></p>
		
<h4>aspl.pCase(value)</h4>
<p><code>aspl.pCase(value)</code> converts value to proper case, in addition to VBScript's lCase and uCase functions.</p>
		
<h4>aspl.curPageURL</h4>
<p><code>aspl.curPageURL</code> returns the url of the current ASP script, including the protocol, server name and script name.</p>
		
<h4>aspl.fso</h4>
<p><code>aspl.fso</code> returns an instance of the VBScript FileSystem Object</p>

<h4>aspl.dict</h4>
<p><code>aspl.dict</code> returns an instance of the VBScript Dictionary Object</p>	

<h4>aspl.recordset</h4>
<p><code>aspl.recordset</code> returns a disconnected adodb.recordset</p>				
				
<h4>aspl.checkEmailSyntax(strEmail)</h4>
<p><code>aspl.checkEmailSyntax</code> checks the syntax of an email address - returns true/false</p>

<h4>aspl.stripHTML(value)</h4>
<p>Strips the HTML tags from value</p>
<p>Example:</p>
<pre class="alert alert-light">&lt;%=aspl.stripHTML("&lt;u&gt;not underlined&lt;/u&gt;")%&gt;</pre>
<p>returns:</p> 
<pre class="alert alert-light"><%=aspl.stripHTML("<u>not underlined</u>")%></pre>

<h4>aspl.padLeft(value, totalLength, paddingChar)</h4>
<p>Adds paddingChar left to value until totalLength is reached.</p>
<p>Example:</p>
<pre class="alert alert-light">&lt;%=aspl.padLeft("5",3,"0")%&gt;</pre>
<p>returns:</p>
<pre class="alert alert-light"><%=aspl.padLeft("5",3,"0")%></pre>


<h4>aspl.padRight(value, totalLength, paddingChar)</h4>
<p>Adds paddingChar right to value until totalLength is reached.</p>
<p>Example:</p>
<pre class="alert alert-light">&lt;%=aspl.padRight("5",3,"0")%&gt;</pre>
<p>returns:</p>
<pre class="alert alert-light"><%=aspl.padRight("5",3,"0")%></pre>

<h4>aspl.getFileType(filename)</h4>
<p>Returns the file extension for a given filename.</p>
<p>Example:</p>
<pre class="alert alert-light">&lt;%=aspl.getFileType("photo.jpeg")%&gt;</pre>
<p>returns:</p>
<pre class="alert alert-light"><%=aspl.getFileType("photo.jpeg")%></pre>
		
<h4>aspl.log(value)</h4>
<p>aspLite comes with a basic logger: <code>aspl.log("anything")</code> will write "anything" to aspLite/aspLite.log. The exact time of the logging is included as well. This logging feature is very useful as you can always tell what exactly happens to a variable, or when things go wrong. I often use it to debug certain functions.</p>

<h4>aspl.recycleApplication</h4>
<p>Resets an ASP application by opening, saving and closing the global.asa-file.</p>	

<h4>aspl.randomizer</h4>
<p>Randomizer class with 3 functions:</p>
<ul>
	<li><code>aspl.randomizer.randomText(nmbrchars)</code> returns a random string with a given nmbrchars</li>
	<li><code>aspl.randomizer.randomNumber(start,stop)</code> returns a random number between start and stop</li>
	<li><code>aspl.randomizer.createGUID(length)</code> returns a GUID of a given length</li>				
</ul>
	
<h4>aspl.removeAllCookies</h4>

<p>Removes all cookies</p>

<h4>aspl.printTimer</h4>

<p>Returns the actual execution time of your ASP script.</p>

<h4>aspl.dump(value)</h4>

<p>Sends value to the browser as text/html. aspLite destroys itself and all its plug-ins right after. In case value includes [aspLite_executionTime], the actual execution time will be added to the output as an HTML comment.</p>	

<h4>Application Caching</h4>

<p>ASP comes with a very interesting object, the Application Object. It was designed to store some application-wide values. In most cases, ASP developers only used it very minimally. They typically stored only very few strings and numbers in the Application Object. However, it was designed to be used a lot heavier than that. It can easily store 10000s of values. When pumped with Arrays, the Application Object is nothing less but a very powerful dictionary or collection object. That's why I built a very tiny layer around it in aspLite.</p>

<p><code>aspl.setcache(name,value)</code> stores an array in the Application Object. Not only the name and the value are stored, but also the exact time. Remember the cacheprefix above? It helps not to interfere with other Application variables you may already use in your ASP script.</p>	

<p><code>aspl.getcache(name)</code> returns the Application value for a given name.</p>

<p><code>aspl.getcacheT(name,seconds)</code> returns the Application value for a given name only if it was stored less than x seconds ago.</p>

<p><code>aspl.clearAllCache()</code> clears all Application values for your cacheprefix.</p>

<h3>Text files and binaries</h3>

<p>For years I have assumed that ASP/VBScript was not capable of dealing with (large) binary files (upload, read, write, save, download). At least, that was, without expensive third party COM components. I was wrong all this time. Lewis Moten once created a purely scripted (and free-to-use) ASP/VBScript Upload class. It did and still does a very good job.</p>

<div class="mt-4 mb-4 p-4  text-bg-info">
<h4>FileSystemObject vs ADODB.Stream</h4>
<p>While the VBScript FileSystemObject is needed to browse files and folders, aspLite uses <strong>ADODB.Stream</strong> to read, write and save files, both binaries and texts. The FileSystemObject does not support binary files and has major issues with dealing with the UTF-8 (Unicode) charset. ADODB.Stream however, perfectly handles binaries and Unicode characters.</p>
</div>

<h4>aspl.loadText(path)</h4>

<p>Returns the content of a text file. Path is the relative path to a text-file.</p>

<p><code>aspl.loadText(path)</code> is a great help when separating ASP-code from HTML, thus preventing the so-called spaghetti-ASP-way of coding. You would typically want to isolate all html blocks and snippets in a specific folder and dynamically include them in your ASP script when needed. </p>

<p><strong>Example:</strong></p>

<p>In the root of your application, you have "default.asp".</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/loadText.inc")))%></pre>

<p>The folder "html" holds default.inc:</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/defaultreplace.inc")))%></pre>

<p>Congratulations! You just wrote your first code-behind ASP file and your first Classic ASP MasterPage.</p>

<div class="mt-4 mb-4 p-4  text-bg-warning">
<h5>Which file-extension should I use for protected and/or code-behind files?</h5>
<p>.inc is a good choice. I also use .resx. Both file types are not served by IIS. In other words, they will never be executed by IIS and they will never be sent nor exposed to the browser. Many other formats will either be sent or executed: .asp, .htm(l), .txt. Be careful with them.</p>
</div>

<h4>aspl.saveText(path,data)</h4>
<p>Saves text (data) to a file (relative path). Unlike the FileSystemObject, this one is 100% UTF-8-proof.</p>

<pre class="alert alert-light">
&lt;!-- #include file="aspLite/aspLite.asp"--&gt;
&lt;%
aspl.saveText "html/textfile.txt", "This is the content of the file."
%&gt;
</pre>

<h4>aspl.saveAsFile(fileName, fileBody)</h4>

<p>Sends text-files (only) to the browser (with a download-prompt). Filename is the default filename, filebody is the content of the file.</p>

<p>At first sight, this is not very useful, as you can send any text to the browser already. However, in some cases, you may want to send text-files (.txt, .html, .css, .js, etc), rather than just text. After showing the download-prompt, they will then end up in the visitor's default download-folder.</p>

<pre class="alert alert-light">
&lt;!-- #include file="aspLite/aspLite.asp"--&gt;
&lt;%
aspl.saveAsFile "textfile.txt", "This is the content of the file."
%&gt;
</pre>

<h4>aspl.loadBinary(filename)</h4>

<p>Returns the content of a binary file (an image, pdf, zip or any other non-text file). Filename is the relative path to a binary file.</p>

<p>As Classic ASP/VBScript does not have many file-editing options, there is not much use for loading binary files, other than using a lot of RAM in case of large files.</p>

<p>Except perhaps when combined with another aspLite function: </p>

<h4>aspl.saveBinaryData(filename,bytearray)</h4>
<p>Saves a given binary file to filename (absolute path!)</p>
<p>The example below copies html/sample.jpg to html/copy.jpg.</p>

<pre class="alert alert-light">
&lt;!-- #include file="aspLite/aspLite.asp"--&gt;
&lt;%
dim file : file=aspl.loadBinary("html/sample.jpg")

aspl.saveBinaryData server.mappath("html/copy.jpg"), file
%&gt;
</pre>

<p>Another use I can think of is to copy a file from anywhere to your local drive. The example below downloads (GET) the JPG file on <a href="https://demo.aspLite.com/default/html/smallfile.jpg" target="_blank">https://demo.aspLite.com/default/html/smallfile.jpg</a> and saves it to your local html folder.</p>

<pre class="alert alert-light">
&lt;!-- #include file="aspLite/aspLite.asp"--&gt;
&lt;%
dim oXMLHTTP : set oXMLHTTP = Server.CreateObject("MsXML2.ServerXMLHTTP")
oXMLHTTP.open "GET", "https://demo.aspLite.com/default/html/smallfile.jpg"
oXMLHTTP.send

aspl.SaveBinaryData server.mappath("html/smallfile.jpg"),oXMLHTTP.responseBody
%&gt;
</pre>

<h4>aspl.dumpBinary (path,dumpAs)</h4>
<p>Sends a binary file (relative path) to the browser, asking the user to download the given file as "dumpAs". dumpAs would typically be the filename with the correct extension (jpg,zip,pdf, etc). </p>

<p>The example below offers a picture of a Suzuki Samurai for download. </p>

<pre class="alert alert-light">
&lt;!-- #include file="aspLite/aspLite.asp"--&gt;
&lt;%
aspl.dumpBinary "html/sample.jpg", "Suzuki Samurai.jpg"
%&gt;
</pre>

<h4>aspl.pathinfo</h4>
<p>Returns custom paths in case a 404-redirect is set up in web.config.</p>

<p>For this to work fine, you need the following web.config file in the root of your application:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/webconfig.inc")))%></pre>

<p>This tells IIS to execute /default.asp in case a 404-error (file not found) is thrown by IIS. Luckily, the missing filename is passed on via Request.ServerVariables("query_string"). That's where <code>aspLite.pathinfo</code> grabs the missing file or folder name.</p>

<p>The aspLite demo site comes with some examples:<br><a href="https://demo.aspLite.com/about">https://demo.aspLite.com/about</a> points to a non-existing folder. <code>aspl.pathinfo</code> returns "about" in this case. From there you can easily serve specific content for the "about" page.</p>

<p>This technique of executing scripts in case of a 404 is available from Windows 2000 (IIS5) onwards. It still works on Windows 2022 servers. But it was never documented or approved by MicroSoft as a valid method. That's a shame, because all this time Classic ASP developers could have made use of this technique to offer user-friendly and SEO-optimized urls. I used this technique in QuickerSite from day one. This has meant a big deal and it has drastically improved SEO for all my hosted Classic ASP customers.</p>

<div class="mt-4 mb-4 p-4  text-bg-dark">
This technique of rewriting 404 urls to "default.asp" could have led, back in 2000, to a true Classic ASP MVC framework. But that didn't happen. It's only after the release of ASP.NET MVC in 2007(!) that some Classic ASP developers came up with MVC-alike ideas for Classic ASP. But by then, Classic ASP developers were way too often ignored and ridiculed by MicroSoft developers. So nothing really happened, again.</div>

<h3>MSXML2.ServerXMLHTTP and MSXML2.DOMDocument</h3>

<p>ASP ships with some very useful http-callers. In short: from within your ASP/VBScript application, you can call basically any other url on the internet, post to any form out there, next wait for its response, and use that output in your application. The output can be text/html, XML or again, pure binary content. I have often used this technique to read RSS-files (typically done with JavaScript today) and copy entire folders recursively from one server to the other (or to any localhost). These ASP functions are also often used to synchronize data between applications via web services.</p>

<p>Another way to look at these two functions is as an Emergency Exit for ASP/VBScript developers. They can be used to organize excursions out of ASP/VBScript and implement PHP or .NET functionalities that are not (easily) available for Classic ASP developers. The technique I'm referring to is to load specific PHP or ASP.NET pages and have them supercharge your ASP script (create images, zip files, pdf files, etc).</p>

<h4>aspl.XMLhttp(url,binary)</h4>
<p>Returns the output of any url - whether or not binary (true/false).</p>
<p>Let's revisit one of the above examples. The line below downloads a JPG file from demo.aspLite.com and saves it to your local html folder. All in one line of ASP code.</p>

<pre class="alert alert-light">
&lt;!-- #include file="aspLite/aspLite.asp"--&gt;
&lt;%
dim file
file="https://demo.aspLite.com/default/html/smallfile.jpg"
aspl.SaveBinaryData server.mappath("html/smallfile.jpg"),aspl.XMLhttp(file,true)
%&gt;
</pre>

<h4>aspl.XMLdom(url)</h4>
<p>Returns an XML document (url) you can list all elements for. Today's web developers would prefer JSON data to XML. More about JSON and aspLite later.</p>

<div class="mt-4 mb-4 p-4  text-bg-warning">

<p><strong>Be careful, XMLhttp and XMLdom perform synchronous requests</strong></p>
<p>Unlike AJAX calls in browsers, these two server-side http-request calls perform synchronous calls. That means that ASP waits for the output of the loaded url before it continues to execute the code below. Therefore, you have to be careful not to load urls that take a long time to execute. That would cause your ASP page to hang.</p>

</div>

<p>An example. The code below shows the 6 latest news items from NYTimes' world-stories-RSS: 
<a href="https://rss.nytimes.com/services/xml/rss/nyt/World.xml" target="_blank">https://rss.nytimes.com/services/xml/rss/nyt/World.xml</a>. The RSS is loaded via <strong>aspL.XMLdom</strong>.</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/feed.inc")))%></pre>

<p><strong>Live example:</strong></p>

<div class="alert alert-light">
<%aspl("ebook/feed.inc")%>
</div>

<div class="mt-4 mb-4 p-4  text-bg-success">
<p><strong>Summary</strong></p>

<p>aspLite comes with quite some powerful functions to deal with text files, binaries, whether they reside on your server or anywhere on the internet. Unlike when using VBScript's FileSystemObject, aspLite ensures that any text converts UTF-8-proof.</p>

<p>Classic ASP's serverhttp-requests can open up a whole new world when integrating with other web-technologies and services.</p>

</div>

<div class="pagebreak"></div>
<a name="chapter5"></a>
<h2>How aspLite lead to asplForm</h2>

<p>When I was developing aspLite during the first weeks of the Covid19 lockdown, I quickly realized aspLite would soon lead to an AJAX formbuilder.</p>

<h3>asplForm explained</h3>

<p>So far aspLite did not rely on any JavaScript framework. And if you don't plan to use asplForm, you don't need jQuery nor any other JavaScript to use aspLite.</p>

<p>asplForm was the first plugin I developed for aspLite. asplForm needs aspLite.js in the browser. aspLite.js relies on jQuery. That makes aspLite a full stack developer framework. Classic ASP developers always were full stack anyway. They needed HTML, CSS and JavaScript for the front and ASP, VBScript, SQL in the back. On top of that, most Classic ASP developers had to deal with IIS administration, SQL Server management, setting up ftp-servers, mail servers, firewalls, backup solutions, Windows server security, etc. That's what I like a lot about being a Classic ASP developer. It's a colorful job with a lot of diversity and complexity. Classic ASP developers were not only full stack, they were often one man bands, taking care of basically everything that related to building applications for their customers.</p>

<div class="mt-4 mb-4 p-4  text-bg-success">

<p><strong>What is asplForm exactly?</strong></p>

<p>asplForm facilitates AJAX web forms. asplForm collects form-fields (VBScript dictionaries), sends them to the browser in JSON-format, and aspLite.js converts it to HTML5 web forms. asplForms bind to any HTML element, regular DIV's in most cases. This technique is the heart of the aspLite framework. Again, asplForm is not only relying on pure Classic ASP/VBScript. There is a very important role for JavaScript when converting server-generated JSON to HTML controls.</p>

</div>

<h4>aspl.form</h4>
<p>Returns an instance of an asplForm. Example: <code>dim form : set form=aspl.form</code>

<h4>form.write(value)</h4>
<p>Adds value as plain text/html to the JSON output of the form.</p>

<h4>form.writejs(value)</h4>
<p>Adds value as JavaScript to the JSON output of the form.</p>

<h4>form.build</h4>
<p>Builds the form, including all its fields and next stops executing the ASP script.</p>

<p>An example. Create a file named <strong>default.asp</strong> in the root of your aspLite application:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/defaultfirst.inc")))%></pre>

<p>If everything goes well, you end up with something like this in your browser:</p>

<pre class="alert alert-light">
This is the header
This is the article
This is the footer
</pre>

<p>When a browser loads default.asp for the first time, it basically loads all HTML, CSS and JavaScript. As soon as they're loaded, asplite.js kicks in. Asplite.js will perform AJAX calls for all HTML elements that have the class "asplForm". These AJAX calls look like this: <strong>default.asp?asplEvent=myheader, default.asp?asplEvent=myarticle and default.asp?asplEvent=myfooter</strong>. Default.asp deals with these "events" in the select case-statement.</p>

<p>Default.asp is loaded 4 times in total. The first time to load all HTML, CSS and JavaScript, the next 3 times to load specific content for specific (onload) events.</p>

<p>Let's have a closer look at the JSON generated by <strong>default.asp?asplEvent=myheader</strong>:</p>

<pre class="alert alert-light">
{
   "target":"myheader",
   "offSet":150,
   "className":"",
   "doScroll":false,
   "id":"myheader_aspForm",
   "onSubmit":"aspAjax('POST',aspLiteAjaxHandler,$(this).serialize(),aspForm);return false;",
   "executionTime":"31ms",
   "bShowToasts":false,
   "aspForm":[
      {
         "type":"hidden",
         "name":"asplPostBack",
         "value":"true",
         "noinit":"true"
      },
      {
         "type":"hidden",
         "name":"asplSessionId",
         "value":"479394439",
         "noinit":"true"
      },
      {
         "type":"hidden",
         "name":"asplTarget",
         "value":"myheader",
         "noinit":"true"
      },
      {
         "type":"hidden",
         "name":"asplEvent",
         "value":"myheader",
         "noinit":"true"
      },
      {
         "type":"plain",
         "html":"this is the header"
      }
   ]
}
</pre>

<ul>
	<li><code>target</code>: tells which HTML-control has load the asplForm</li>
	<li><code>offSet</code>: currently not used</li>
	<li><code>className</code>: the CSS-class for the form-tag of the asplForm</li>
	<li><code>doScroll</code>: currently not used</li>
	<li><code>id</code>: the id of the form</li>
	<li><code>onSubmit</code>: the JavaScript to be executed when a form is submitted</li>
	<li><code>executionTime</code>: the time (ms) it took to process the asplForm and generate the JSON</li>
	<li><code>bShowToasts</code>: currently not used</li>
	<li><code>aspForm</code>: a JSON object including all the form-fields</li>
	<li><code>aspForm.asplPostBack</code>: always true, telling asplForm that a form was submitted</li>
	<li><code>aspForm.asplSessionId</code>: ASP SessionID, used to avoid session-riding, also known as Cross-Site Request Forgery (CSRF)</li>
	<li><code>aspForm.asplTarget</code>: target-div for asplForm</li>
	<li><code>aspForm.asplEvent</code>: repeats the incoming asplEvent for asplForm to know which code to run</li>	
</ul>

<p>The <code>form.build</code>-method stops the execution of the ASP script. That's why it's very important to call it here. If we don't, it would each time add the HTML, CSS and JavaScript to its response. We already have these, so we don't need them again.</p>

<p>Some might think: is it a good idea to load default.asp 4 times (in a row)? Maybe not. Or maybe it doesn't really matter. The most important idea however is that from now on, it's possible to build a web application based on events and clearly organize your ASP code base, using one ASP entry-point (default.asp) only, performing an unlimited number of different tasks. How you structure it is up to you. There are several ways to do it. You can even use good old include-files for each asplEvent, or use aspLite's execute methods. </p>

<p><strong>Let's take this one step further</strong></p>

<p>Use this <strong>default.asp</strong> in the root of your ASP application:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/default.inc")))%></pre>

<p>Note that we're not loading (including) one particular file. We're loading "whatever" value <code>aspl.getRequest("asplEvent")</code> holds. This way, we dynamically load ASP code into the ASP page's namespace. If a file does not exist, aspLite throws a file-not-found error in case aspL_debug (aspLite/config.asp) is "true".</p>

<p>Next, create <strong>html/default.inc</strong> (create the HTML-folder if needed) with the following HTML, CSS and JavaScript:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/defaulthtml.inc")))%></pre>

<p>Finally, create the asp folder and add 3 files to it:</p>

<p><strong>asp/myheader.inc:</strong></p>
<pre class="alert alert-light">
&lt;%
dim form : set form=aspl.form
form.write "this is the header"
form.build
%&gt;
</pre>

<p><strong>asp/myarticle.inc:</strong></p>

<pre class="alert alert-light">
&lt;%
dim form : set form=aspl.form
form.reload=1
form.write "this is the servertime: " & now()
form.build
%&gt;
</pre>

<p><strong>asp/myfooter.inc:</strong></p>
<pre class="alert alert-light">
&lt;%
dim form : set form=aspl.form
form.write "this is the footer"
form.build
%&gt;
</pre>

<p>Browsing to default.asp will return the following HTML:</p>

<pre class="alert alert-light">
this is the header
this is the servertime: 5-4-2024 18:25:08
this is the footer
</pre>

<p>A <strong>clock</strong> starts to run in <strong>asp/myarticle.inc</strong>. As <code>form.reload</code> was set to 1 for myarticle.inc, the asplForm in myarticle.inc gets reloaded every second.</p>

<p>Also, this time the code is organized in a different way. Call it code-behind or anything you like. Default.asp executes ASP code, html/default.inc holds all HTML, CSS and JavaScript and the 3 asp-includes each return an asplForm with some content.</p>

<div class="mt-4 mb-4 p-4  text-bg-info">
<p><strong>Why would you reload an aspLite form every x seconds?</strong></p>
<ul>
	<li>To keep an ASP session alive.</li>
	<li>To list all currently logged on users</li>
	<li>To offer real time game-scores</li>
	<li>...</li>
</ul>
</div>

<p><strong>Can't I just use include-files instead of using aspl.exec()?</strong></p>
<p>Sure. This would have worked as well:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/includes.inc")))%></pre>

<p>Be careful though! I moved the line <code>dim form : set form=aspl.form</code> to default.asp and removed it in all 3 *.inc files. If not, we get a "name redefined" error. That error is not thrown when using <code>aspl.exec()</code> - or VBScript's executeGlobal. Not sure why that is the case. Maybe it's a bug in executeGlobal.</p>

<p>As such, "good old" including all different ASP scripts like this will not harm performance nor will it need much more RAM. The only inconvenience is that you have to list all events and "manually" assign the appropriate script to them, like in the example above. There is one other reason why you'd prefer regular includes: better error descriptions! Whenever something goes wrong in ASP code imported by <code>aspl.exec()</code>, all you get is a vague error message without a line number! That can be very annoying to debug your ASP code.</p>

<h4>form.field(fieldType)</h4>

<p>Returns a VBScript Dictionary. Example: <code>dim field : set field=form.field("select")</code> returns an HTML selectbox.</p>

<p>So far we've used asplForms to serve plain text/html. There is much more to HTML forms than plain text/html. These are the various field-types for asplForms:</p>

<ul>
	<li>form.field("hidden")</li>
	<li>form.field("text")</li>
	<li>form.field("date")</li>
	<li>form.field("email")</li>
	<li>form.field("number")</li>
	<li>form.field("select")</li>
	<li>form.field("textarea")</li>
	<li>form.field("radio")</li>
	<li>form.field("checkbox")</li>	
	<li>form.field("button")</li>
	<li>form.field("submit")</li>
	<li>form.field("reset")</li>
	<li>form.field("plain")</li>
	<li>form.field("script")</li>
</ul>

<p>Also, it's about time to import a true HTML & CSS framework. I personally like Bootstrap a lot. At this time of writing, Bootstrap 5.3 is the latest version and considered the most popular front-end development framework.</p>

<p>Whenever you start a new ASP development, always use a modern front-end framework. Do not write your own, use an existing one. Back in the late nineties, ASP developers were used to hand code HTML controls, very often on an HTML table-layout. You could really tell if a website was on Classic ASP/VBScript back in those days. They all looked very similar.</p>

<p>An average customer never cares about which technologies their website uses in the back. Customers only care about what they see in the front: an attractive design, good usability, nice colours, does it work on a phone, the icon-set you use, etc. Your application can use the very latest ASP.NET version, if it doesn't have an attractive font-end, your customer won't be happy. On the other hand, if you're using a state-of-the-art front-end framework and well-coded Classic ASP/VBScript in the back, your invoices get paid with a smile.</p>

<p><strong>An example</strong></p>

<p>Let's reuse the default.asp from above in the root:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/default.inc")))%></pre>

<p>Add <strong>html/default.inc</strong> and import the Bootstrap CSS and JS libraries via CDN:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/bootstrap.inc")))%></pre>

<p>Finally, add <strong>asp/main.inc</strong>:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/fields.inc")))%></pre>

<p>As <code>form.field("")</code> returns a VBScript dictionary, the following syntax works as well:</p>

<pre class="alert alert-light">
dim hidden : set hidden=form.field("hidden")
hidden("name")="hidden"
hidden("value")="12345"
</pre>

<p>If all goes well, you end up with a form asking for your name, your birth date, your email address and some remarks. After you successfully submit, some confirmation message shows up. Congratulations! This is your first (fully responsive) AJAX form using Bootstrap, aspLite and asplForm!
</p>

<p><strong>Live example:</strong></p>
<div class="asplForm alert alert-light" id="fields"></div>

<p>A lot is happening in the back:</p>

<ol>

	<li>The browser requests <strong>default.asp</strong></li>
	<li>As there is <strong>no asplEvent</strong> yet, default.asp loads <strong>html/default.inc</strong>, our HTML, CSS and JavaScript front-end</li>
	<li>Asplite.js initializes all HTML-controls that have the class "asplform" through <strong>AJAX calls</strong>, only one in this case: default.asp?asplEvent=main</li>
	<li>The server loads <strong>/asp/main.inc</strong> - our asplForm - in case default.asp?asplEvent=main is requested</li>
	<li>Asplite.js builds an <strong>HTML form</strong> with all elements that are returned by default.asp?asplEvent=main</li>
	<li>When hitting the submit-button, the form is submitted through an AJAX-call and form.postback will be true</li>
	<li>If form.postback is true, you can process all values coming from the form.</li>

</ol>

<p><strong>default.asp?asplEvent=main</strong> returns <strong>JSON</strong>:</p>

<pre class="alert alert-light">
{
   "target":"main",
   "offSet":150,
   "className":"",
   "doScroll":false,
   "id":"main_aspForm",
   "onSubmit":"aspAjax('POST',aspLiteAjaxHandler,$(this).serialize(),aspForm);return false;",
   "executionTime":"31ms",
   "bShowToasts":false,
   "aspForm":[
      {
         "type":"hidden",
         "name":"asplPostBack",
         "value":"true",
         "noinit":"true"
      },
      {
         "type":"hidden",
         "name":"asplSessionId",
         "value":"575517568",
         "noinit":"true"
      },
      {
         "type":"hidden",
         "name":"asplTarget",
         "value":"main",
         "noinit":"true"
      },
      {
         "type":"hidden",
         "name":"asplEvent",
         "value":"main",
         "noinit":"true"
      },
      {
         "type":"hidden",
         "name":"hidden",
         "value":"12345"
      },
      {
         "type":"text",
         "name":"name",
         "class":"form-control",
         "label":"Your name",
         "required":true
      },
      {
         "type":"plain",
         "html":"&lt;div style=\u0022clear:both;height:7px\u0022 class=\u0022clearfix\u0022&gt;&lt;\u002Fdiv&gt;"
      },
      {
         "type":"number",
         "name":"years",
         "class":"form-control",
         "label":"For how many years are you an ASP developer?"
      },
      {
         "type":"plain",
         "html":"&lt;div style=\u0022clear:both;height:7px\u0022 class=\u0022clearfix\u0022&gt;&lt;\u002Fdiv&gt;"
      },
      {
         "type":"date",
         "name":"birthdate",
         "class":"form-control",
         "label":"Your birthdate"
      },
      {
         "type":"plain",
         "html":"&lt;div style=\u0022clear:both;height:7px\u0022 class=\u0022clearfix\u0022&gt;&lt;\u002Fdiv&gt;"
      },
      {
         "type":"email",
         "name":"email",
         "class":"form-control",
         "label":"Your email"
      },
      {
         "type":"plain",
         "html":"&lt;div style=\u0022clear:both;height:7px\u0022 class=\u0022clearfix\u0022&gt;&lt;\u002Fdiv&gt;"
      },
      {
         "type":"textarea",
         "name":"remarks",
         "class":"form-control",
         "label":"Remarks",
         "rows":5
      },
      {
         "type":"plain",
         "html":"&lt;div style=\u0022clear:both;height:7px\u0022 class=\u0022clearfix\u0022&gt;&lt;\u002Fdiv&gt;"
      },
      {
         "type":"submit",
         "html":"Submit",
         "class":"btn btn-primary"
      },
      {
         "type":"reset",
         "html":"Reset",
         "class":"btn btn-secondary"
      }
   ]
}
</pre>

<div class="mt-4 mb-4 p-4  text-bg-primary">
<p>This JSON-stream is generated by the JSON class in asplite/asplite.asp. It's a 100% copy of the JSON class I found in ASP-Ajaxed on <a target="_blank" style="color:#FFF" class="link" href="https://github.com/ASP-Ajaxed/asp-ajaxed">https://github.com/ASP-Ajaxed/asp-ajaxed</a>. It was originally developed by Michal Gabrukiewicz. His passing in 2009 was surely one of the reasons for this framework not to become what it could have been. With aspLite, I want to honour Michal for his great work on ASP-Ajaxed. He was one of the best ASP developers ever. I consider aspLite at least 50% Michal's work.</p>

<p>JSON parsers are amongst the most popular Classic ASP repositories on GitHub.</p>

<ul>
	<li><a class="link" style="color:#FFF" target="_blank" href="https://github.com/rcdmk/aspJSON">aspJSON</a> by Ricardo Souza</li>
	<li><a style="color:#FFF" class="link" href="https://github.com/gerritvankuipers/aspjson" target="_blank">aspjson</a> by Gerrit Van Kuipers</li>
</ul>

<p>This illustrates how developers find their way around limitations and shortcomings. MicroSoft never added support for JSON to Classic ASP. But that does not stop ASP developers. We need JSON-functionality to integrate modern JavaScript libraries and frameworks into our applications. So we get it done!</p>
</div>

<h4>form.newline</h4>
<p>Adds a line feed of 7px (height).</p>

<p>Example with <strong>selectboxes, checkboxes and radio buttons</strong>. Keep both default.asp and html/default.inc. <strong>Change asp/main.inc to:</strong></p>


<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/selectboxes.inc")))%></pre>

<p><strong>Live example:</strong></p>

<div class="asplForm alert alert-light" id="selectboxes"></div>

<p>You can see how a VBScript Dictionary is translated into HTML select boxes, a list of radio buttons and checkboxes by <strong>aspLite.js</strong>. Be careful, only VBScript dictionaries are supported! They are a perfect match with these HTML-controls (key/value-pairs).</p>

<h4>form.submit</h4>
<p>In the example above, the VBScript helper function <code>form.submit</code> returns the JavaScript code necessary to submit the asplForm.</p>

<p>The <strong>JSON string</strong> for <strong>asp/main.inc</strong> looks like this:</p>

<pre class="alert alert-light">
{
   "target":"main",
   "offSet":150,
   "className":"",
   "doScroll":false,
   "id":"main_aspForm",
   "onSubmit":"aspAjax('POST',aspLiteAjaxHandler,$(this).serialize(),aspForm);return false;",
   "executionTime":"55ms",
   "bShowToasts":false,
   "aspForm":[
      {
         "type":"hidden",
         "name":"asplPostBack",
         "value":"true",
         "noinit":"true"
      },
      {
         "type":"hidden",
         "name":"asplSessionId",
         "value":"575517568",
         "noinit":"true"
      },
      {
         "type":"hidden",
         "name":"asplTarget",
         "value":"main",
         "noinit":"true"
      },
      {
         "type":"hidden",
         "name":"asplEvent",
         "value":"main",
         "noinit":"true"
      },
      {
         "type":"select",
         "name":"selectbox",
         "emptyfirst":"",
         "class":"form-control form-select",
         "label":"Selectbox showing all values of a VBScript dictionary",
         "options":{
            "1":"option1",
            "2":"option2",
            "3":"option3",
            "4":"option4",
            "5":"option5",
            "6":"option6",
            "7":"option7",
            "8":"option8",
            "9":"option9"
         },
         "onchange":"$('#main_aspForm').submit();return false;"
      },
      {
         "type":"plain",
         "html":"&lt;div style=\u0022clear:both;height:7px\u0022 class=\u0022clearfix\u0022&gt;&lt;\u002Fdiv&gt;"
      },
      {
         "type":"radio",
         "name":"radio",
         "class":"form-check-input",
         "labelclass":"form-check-label",
         "label":"Radiobutton showing all values of a VBScript dictionary",
         "options":{
            "1":"option1",
            "2":"option2",
            "3":"option3",
            "4":"option4",
            "5":"option5",
            "6":"option6",
            "7":"option7",
            "8":"option8",
            "9":"option9"
         },
         "onchange":"$('#main_aspForm').submit();return false;"
      },
      {
         "type":"plain",
         "html":"&lt;div style=\u0022clear:both;height:7px\u0022 class=\u0022clearfix\u0022&gt;&lt;\u002Fdiv&gt;"
      },
      {
         "type":"plain",
         "html":"Checkboxes showing all values of a VBScript dictionary"
      },
      {
         "type":"checkbox",
         "class":"form-check-input",
         "labelclass":"form-check-label",
         "name":"checkboxes",
         "options":{
            "1":"option1",
            "2":"option2",
            "3":"option3",
            "4":"option4",
            "5":"option5",
            "6":"option6",
            "7":"option7",
            "8":"option8",
            "9":"option9"
         },
         "container":"div",
         "containerclass":"form-check form-switch",
         "onclick":"$('#main_aspForm').submit();return false;"
      }
   ]
}
</pre>

<p>You can see how the various collections of "options" are passed through as JSON objects.</p>

<h4>form.field("plain") and form.field("script")</h4>
<p>Writes pure text/html (plain) and JavaScript (script)</p>
<p>Example: keep default.asp and html/default.inc. <strong>Change asp/main.inc to</strong>: </p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/plain.inc")))%></pre>

<p><strong>Live example:</strong></p>
<div class="asplForm alert alert-light" id="plain"></div>

<p>There are shortcuts for both field-types (plain and script):</p>

<pre class="alert alert-light">
form.write "This adds plain/text or &lt;u&gt;HTML&lt;/u&gt;"
form.writejs "alert('Add JavaScripts');"
</pre>

<div class="mt-4 mb-4 p-4  text-bg-primary">

<p><strong>asplForms come with a VIEWSTATE-kind of facility</strong></p>

<p>By default asplForms keep track of form-values and auto-fills form-controls with user-input. There is no need to initialize them "manually" with request-values.</p>

</div>

<h4>form.alert(alerttype,message)</h4>

<p>This helper-function relies on Bootstrap 5. It returns Bootstrap alerts.</p>

<p>Change asp/main.inc to:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/alert.inc")))%></pre>

<p><strong>Live example:</strong></p>
<div class="asplForm alert alert-light" id="alert"></div>

<h4>More complex forms with Bootstrap</h4>
<p>Sometimes you need more flexibility when it comes to placing form-controls.</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/complex.inc")))%></pre>

<p><strong>Live example:</strong></p>
<div class="asplForm alert alert-light" id="complex"></div>

<p>By setting <code>form.className="row"</code>, we initialize Bootstrap's 12-column grid-system. <code>name.add "containerclass","col-md-6"</code> makes sure the "name" field uses half of the available grid-lines (6). This way you can easily have multiple form-fields on 1 line. The good thing about using Bootstrap for this purpose, is that forms remain 100% responsive.</p>


<h4>form.redirect(asplevent,target,querystring)</h4>
<p>Helper function to navigate between asplForms. asplEvent indicates which ASP code to load, target is the ID of the HTML-control that will be bound to and querystring can be used to add custom parameters.</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/redirect.inc")))%></pre>

<p><strong>Live example:</strong></p>
<div class="asplForm alert alert-light" id="redirect"></div>

<p>After clicking the Redirect-button, the file redirectNew.inc is loaded:</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/redirectNew.inc")))%></pre>

<p>After clicking the back-button, the ASP code in redirect.inc is loaded again. If "target" is empty, it will inherit the name of asplEvent, thus load the script "ebook/redirect.inc" in div-id "redirect".</p>

<p>Unlike when using "response.redirect" in Classic ASP, you can "redirect" to multiple divs. In other words: multiple divs can be reloaded.</p>

<p><strong>ebook/redirectmultiple.inc reads:</strong></p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/redirectmultiple.inc")))%></pre>

<p>After hitting the "Redirect"-button, 3 divs are loaded: redirect1, redirect2 and redirect3. They each load the same ASP code (as they have the same asplEvent: redirectnewmultiple).</p>

<p><strong>ebook/redirectnewmultiple.inc reads:</strong></p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/redirectnewmultiple.inc")))%></pre>

<p><strong>Live example:</strong></p>
<div class="asplForm alert alert-light" id="redirectmultiple"></div>

<div id="redirect1"></div>
<div id="redirect2"></div>
<div id="redirect3"></div>

<div class="mt-4 mb-4 p-4 text-bg-success">

<p><strong>Summary</strong></p>

<p>asplForms turn aspLite into a modern full-stack developer framework for Classic ASP/VBScript developers. It offers a powerful AJAX web form-builder whether or not making use of the wonderful Bootstrap front-end HTML/CSS framework.</p>

<p>The collaboration between server-side Classic ASP/VBScript and browser-side JavaScript frameworks is very promising and opens up a whole new world to ASP development.</p>

<p>All-in-all, this is a very first attempt to bring back Classic ASP/VBScript development. I would even call it "embryonic". Especially the JavaScript file asplite/asplite.js needs some attention and further tweaking.</p>

<p>It should not be too complicated to add even more helper-functions and as aspLite will grow over the years, more helper functions will be needed to facilitate even more exciting features. For now however I prefer to keep things compact, simple and very robust. That's what I liked about Classic ASP in the first place.</p>

<p>The next chapters will cover more interesting topics and will include many more ideas to make Classic ASP development fun again, whether or not using aspLite as the preferred ASP/VBScript framework.</p>

</div>

<div class="pagebreak"></div>
<a name="chapter6"></a>
<h2>aspLite plug-ins</h2>

<p>aspLite comes with some plug-ins. These plug-ins are located in the asplite/plugins-folder. All plug-ins should go in that folder.</p>
<p>There is an important rule to keep in mind when adding your own plug-ins: the name of the folder/ASP-file has to correspond with the name of the actual VBScript Class.</p>
<p>Example: one of the folders in asplite/plugins is "database". The ASP file of the actual plugin is also named "database.asp". That is a requirement for aspLite-plugins to work. Other plug-ins are /asplite/sha256/sha256.asp and /asplite/atom/atom.asp.</p>
<p>Not only the folder and filenames are important, the name of the actual VBScript Class also has to correspond. If you browse the plug-ins, you'll notice the following convention: cls_asplite_database, cls_asplite_atom, cls_asplite_sha256, etc. So there again, the exact same name of the plugin-folder and ASP file returns.</p>
<p>This setup guarantees that aspLite plug-ins work, and all work the same way.</p>
<p><strong>How to create an instance of an aspLite plug-in?</strong></p>

<ul class="list-group">
	<li class="list-group-item"><code>dim db : set db=aspl.plugin("database")</code> creates an instance of the class <code>cls_asplite_database</code>, located in /asplite/plugins/database/database.asp</li>
	<li class="list-group-item"><code>dim sha256 : set db=aspl.plugin("sha256")</code> creates an instance of the class <code>cls_asplite_sha256</code>, located in /asplite/plugins/sha256/sha256.asp</li>
	<li class="list-group-item"><code>dim uploader : set db=aspl.plugin("uploader")</code> creates an instance of the class <code>cls_asplite_uploader</code>, located in /asplite/plugins/uploader/uploader.asp</li>
</ul>

<p>I'm sure you get the point.</p>

<h3>CDO.message</h3>

<p>One of the most robust and best supported mail-sending components from IIS 5 (Windows 2000) onwards, has been the <code>"CDO.Message"</code>-object. Pretty much all (shared) hosting solutions supported it. And it still works on Windows Server 2022. <code>CDO.Message</code> was the successor of <code>"CDONTS.NewMail"</code>. I have been successfully using <code>CDO.Message</code> in QuickerSite for almost 25 years now. Never ran into problems.</p>

<p>aspLite comes with a wrapper-class that facilitates the use of this widely supported ASP component. Be aware that I have commented out all .send() methods to prevent abuse. You would need to un-comment the .send-methods for the scripts to work.</p>

<p>This first example does not use a SMTP-username and password, using your server's own SMTP server:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/cdo.inc")))%></pre>

<p><strong>Live example:</strong></p>
<div class="asplForm alert alert-light" id="cdo"></div>

<p>By default, the mail-body is wrapped into a valid HTML5-template ensuring Unicode-compatibility and correct HTML output. </p>

<p>The next example uses an SMTP-username and password, using an external SMTP server and it adds an attachment. The SMTP server also requires communication via SSL:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/cdo2.inc")))%></pre>

<p>This last example uses a mobile-friendly (responsive) mail-template that I used on <a target="_blank" href="https://Setlistplanner.com">Setlistplanner.com</a>. Feel free to use it for your own projects.</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/cdo3.inc")))%></pre>

<p>The HTML/CSS for <strong>ebook/mail.txt</strong> reads:</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/mail.txt")))%></pre>

<div class="mt-4 mb-4 p-4  text-bg-warning">
<p><strong>Warning: Windows Servers 2025 do no longer come with an SMTP-server!</strong></p>
<p>Windows Servers always shipped with a built-in SMTP server. When I recently gave Windows Server 2022 a try, it appeared to me that the built-in SMTP server was still there, but it got broken. It was no surprise to read that from Windows Server 2025 onwards, the built-in SMTP service will be removed, as it was part of IIS6. IIS6 will also no longer be available in Windows Server 2025.</p>

<p>You may need a replacement for the built-in SMTP server. I always use the freeware version of MailEnable. However, be aware that nowadays the vast majority of email-providers require domain-specific SPF-records that point to the sending server(s). This is something to keep in mind when setting up email-sending applications for your customers. To avoid spam as much as possible, you can better send email over the appropriate mail servers for a given domain. Always check the MX-records for your domain on websites like <a style="color:#000" target="_blank" href="https://mxtoolbox.com/SuperTool.aspx">https://mxtoolbox.com/SuperTool.aspx</a></p>

</div>

<h3>Database</h3>

<p>The most important reason for using VBScript back in 1997-2002 was ADO (ActiveX Data Objects). ADO is a programming interface to access data in a database. When combined with ODBC (Open DataBase Connectivity), Classic ASP/VBScript developers have always been able to connect to Oracle-databases, PostgreSQL, MySQL, Microsoft SQL Server, Access, Excel, DB2, etc. The sky's the limit.</p>

<p>The only two ADO-objects I'ever used in each and every ASP application, are the <code>ADODB.Connection</code> and the <code>ADODB.Recordset</code>. I actually use 2 types of recordsets: read-only (adLockReadOnly - the default lock-type) and read/write (adLockOptimistic). They're all you'll ever need.</p>

<p>aspLite comes with a plugin and wrapper-class for <code>ADODB.Connection</code> and <code>ADODB.Recordset</code>:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/db.inc")))%></pre>

<p><strong>Live example:</strong></p>
<div class="asplForm alert alert-light" id="db"></div>

<p>Every time this ebook is loaded, a "person" gets created in the database. It's an easy trick for me to see how often this book is read. I'm showing only 10 records though.</p>

<p>This little script is doing a lot in the back:</p>

<ul>
<li><code>dim db : set db=aspL.plugin("database")</code> creates an instance of the database-plugin</li>
<li><code>db.dbms=1</code> specifies we're dealing with an Access database</li>
<li><code>db.path="ebook/access.mdb"</code> sets the path to the database</li>
<li><code>dim rs : set rs=db.rs</code> creates a read/write recordset, next a record is added to the table "person"</li>
<li><code>set rs=db.execute("select top 10 * from person")</code> creates a read-only recordset and next loops through the records and sends them to the browser.</li>
</ul>

<p>The database-plugin in aspLite is designed to open a connection to a database <i>only once</i> through the lifespan of an ASP script/page. Opening and closing a connection to a database is time-consuming. You can therefore better avoid opening and closing database connections more than once in your ASP script. In other words: you can better create a database-plugin <strong><i>only once</i></strong> through the lifespan of your ASP script. This is very important to keep the speed and performance of your ASP application optimal.</p>

<p>Please note that we're using an Access database here, like we do in the aspLite demosite. In case you mainly use read-only recordsets and you use read/write recordsets with care, Access databases are your best friend, even for large and heavily used datasets.</p>

<div class="alert alert-info">

	<p><strong>Access databases?</strong></p>

	<p>A lot has been said about using Access databases in web applications (whether or not developed in Classic ASP). Bottom line, most experts say: <em>don't</em>! However, I say: Access is the easiest, fastest, most server-friendly (uses <strong>no RAM</strong>!) and an extremely reliable database to use with Classic ASP. Period. I have been using them for 25 years now. Never ran into a single problem. And today - with these fast SSD drives and powerful CPUs - you're not going to crash an ASP page that's using an Access database. I have gone through all the available stress-tests over the years. I was even getting frustrated at some point, as I have invested a lot of time and money in migrating to SQL Server in the past. All in vain. Waste of time and money. I went back to Access for all my hosted websites in 2018. No regrets.</p>

	<p>That said, there are some guidelines and things to keep in mind when using Access in Classic ASP:</p>
	
	<ul><li><u>Always use <strong>read-only recordsets</strong></u> (they're read-only by default), unless you really need to update records.</li><li><u>Do not store</u> <strong>binary files</strong> (images, pdf, documents) in Access databases. Store them as regular files. If you're concerned about security, give your files a secure (not downloadable) extension (like .resx) and stream them to users with <em>adodb.stream</em>. aspLite facilitates this with its <em>dumpBinary </em></li><li><u>Do not store</u> <strong>visitor data</strong> (logfiles) in a database. Again, use the file system. There is no point in storing bulk data that you're probably never going to look into afterwards anyway. Also, visitor data are typically stored in the IIS log files already. No need to duplicate them.</li><li><u>Do not store</u> <strong>data-backups</strong> in the database itself. Some developers log each and every change to database-records in yet another (copy-)table in order to keep track of changes. In some cases that could be useful. But when using Access, it's very easy to keep a couple of backups to revert to or look into in case data may have been lost or corrupted. Again. Do not store bulk data in an Access database if you're not gonna need it. Make backups... I have a customer on a 250MB Access database. Manually making a copy of that file takes... <em>half a second</em> on a fast SSD drive. So there is no reason to not make Access backups every day, or even every couple of hours.&nbsp;</li><li>Summarizing: <em><strong>only store text and numeric data</strong></em> in Access databases that you're actually going to <em><strong>use</strong> </em>in your application: read, update, delete and search for.</li><li>Make <strong>backups</strong> of your Access mdb files. Do it. Every day, or even every couple of hours for business critical applications. You really don't want to lose data.</li><li>Have a look at the <strong>database-plugin</strong>&nbsp;: /aspLite/plugins/database/database.asp. Ideally, an ASP page opens a connection to a database only ONCE through its lifespan. Opening (and closing) database connections are probably the most time-consuming operations in ASP pages. Doing this only once drastically speeds up your ASP pages. </li><li>There is a limit of <strong>255 concurrent users</strong> for Access databases. However, when using Access in a web application, the "user" (IUSR in most cases) is only connected for a few milliseconds. You will not often face situations where 255 visitors simultaneously open an ASP page <em>(that is: on the very same millisecond)</em>. Unless you're Google or Facebook... I have never faced that situation in 20 years.</li><li>There is a (hard!) limit of <strong>2GB </strong>for Access databases. That's a lot of text and numeric data. Do not let Access databases grow <em>even close</em> to 2GB. I would personally decide to upsize to SQL Server as soon as an Access mdb file grows bigger than 500MB. But that's just me, given the hardware I use and trust (EC2 instance on AWS). I currently do not host a single Access database bigger than 350MB.</li></ul>

</div>

<h3>Sha256 (Secure Hash Algorithm)</h3>

<p>The sha256 ASP/VBScript code was originally written by Phil Fresle in 2001. It still works like a charm, after all these years.</p>

<p>The sha256 plugin would typically be used to store user-passwords in a database. Passwords that are "hashed" with sha256 are one-way encrypted. There is no way to restore passwords once they are hashed with sha256.</p>

<p>Example:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/sha256.inc")))%></pre>

<p><strong>Live example:</strong></p>
<div class="asplForm alert alert-light" id="sha256"></div>

<h3>Jpeg</h3>

<div class="alert alert-warning">This plugin needs ASP.NET 2.0 (or higher) enabled in IIS. </div>

<p>The Jpeg plugin can be used to resize, recolor and crop jpg/png-files. This is unobtrusive. The <a href="/default/html/smallfile.jpg" target="_blank">original jpg-file</a> is not being modified in the process. This plugin works fine for images up to 1MB. When the JPG files are (much) bigger, converting may take some (processing) time and hurts your server's processor.</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/jpeg.inc")))%></pre>

<p><strong>Live example:</strong></p>

<div class="asplForm alert alert-light" id="jpeg"></div>

<h3>Uploader</h3>

<p>Before I started using Lewis Moten's pure ASP/VBScript upload-class, I had used 3rd party Upload-COM components. Back then, the most known company (by far) for developing these components was Persits. They built (and still do) a variety of COM-products for both Classic ASP and ASP.NET.</p>

<p>Uploading files in Classic ASP/VBScript applications has always been very challenging. That was because ASP's Request.form-method did not work when a form was submitted with binary data (enctype="multipart/form-data"). I have always looked at this as one a the most annoying bugs in Classic ASP. And it never got fixed.</p>

<p>Over the years I learned to live with this limitation and I was able to build various fileupload-scripts, whether or not using AJAX to smoothen the user-experience. At some point I even started using a JavaScript function that resized JPG-files before they were uploaded. Very useful in case you needed customers to upload lots of photos, but you wanted to speed-up the upload-process. The demo of aspLite ships with such an upload form. Check it out!</p>

<p>I amended Lewis Moten's code so that only a limited number of filetypes can be uploaded. I didn't want users to upload ASP(x) or PHP scripts. The following file-types are allowed for upload: "jpg", "jpeg", "jpe", "jp2", "jfif", "gif", "bmp", "png", "psd", "eps", "ico", "tif", "tiff", "ai", "raw", "tga", "mng", "svg", "doc", "rtf", "txt", "wpd", "wps", "csv", "xml", "xsd", "sql", "pdf", "xls", "mdb", "ppt", "docx", "xlsx", "pptx", "ppsx", "artx", "mp3", "wma", "mid", "midi", "mp4", "mpg", "mpeg", "wav", "ram", "ra", "avi", "mov", "flv", "m4a", "m4v", "htm", "html", "css", "swf", "js", "rar", "zip", "ogv", "ogg", "webm", "tar", "gz", "eot", "ttf", "ics", "woff", "cod", "msg", "odt". I agree this is quite an arbitrary list, but so far it did the trick for me.</p>

<p>The example for this ebook uploads one file at a time. That can easily be updated to "multiple" files, simply by adding the attribute <strong>"multiple"</strong> to the HTML-fileupload control in <strong>ebook/uploadhtml.inc</strong>. The upload process may seem a little complex, but it is an AJAX upload-facility after all, giving a nice user-experience without annoying page-reloads during the entire process.</p>

<p>Various files are used:</p>

<p>First, the browser requests <strong>/ebook/upload.inc:</strong></p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/upload.inc")))%></pre>

<p>As the actual upload will be performed by a custom JavaScript, I disable aspLite's default "onsubmit" by emptying <code>form.onSubmit=""</code></p>

<p>Furthermore, the ASP script loads 2 files:</p>
<ul>
	<li>one HTML snippet (holding the HTML fileupload-control)</li>
	<li>one JavaScript file (dealing with the actual upload):</li>
</ul>

<p>The HTML snippet in <strong>ebook/uploadhtml.inc</strong>:</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/uploadhtml.inc")))%></pre>

<p><i>Tip: if you add the attribute "multiple" to the HTML-control above, you allow/enable uploading multiple files.</i></p>

<p>The JavaScript in <strong>ebook/uploadjs.inc</strong>:</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/uploadjs.inc")))%></pre>

<p>The actual file-saving takes place in <strong>ebook/uploadfile.inc</strong> (note that the actual "save"-method is commented out - no files are saved for security reasons!). You would need to uncomment that line to enable the actual file-saving.</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/uploadfile.inc")))%></pre>

<p>The JavaScript script itself loads a final file, <strong>ebook/uploadfb.inc</strong> after a successful upload:</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/uploadfb.inc")))%></pre>

<p><strong>Live example:</strong></p>
<div class="asplForm alert alert-light" id="upload"></div>

<p>All in all, the upload-demo in <a href="https://demo.asplite.com" target="_blank">https://demo.asplite.com</a> is very similar, but adds some very useful feedback-messages for each uploaded file and also summarizes the final result.</p>

<h3>Create your own</h3>

<p>The obligate "/asplite/helloworld/helloworld.asp" plug-in demonstrates how to create your own . The most important rule to keep in mind: <strong>never response.write anything to the browser</strong> from within a plugin. aspLite and aspForm rely on JSON-streams that talk to aspLite.js so they can be converted to HTML. If you expect any output from a plugin, use public functions or variables to expose them to the calling script.</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("asplite/plugins/helloworld/helloworld.asp")))%></pre>

<p>The plugin can be used like this:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/helloworld.inc")))%></pre>

<p><strong>Live example:</strong></p>
<div class="asplForm alert alert-light" id="helloworld"></div>

<p>Custom plug-ins are not limited to asplForms though. You can also directly use them in a more classic way:</p>

<pre class="alert alert-light">&lt;%=aspl.plugin("helloworld").hw%&gt;</pre>

<p><strong>Live example:</strong></p>
<div class="alert alert-light"><%=aspl.plugin("helloworld").hw%></div>

<div class="mt-4 mb-4 p-4  text-bg-success">

<p><strong>Summarizing</strong></p>

<p>aspLite comes with a couple of plug-ins (VBScript classes) to facilitate emailing, database, rss & atom, file-upload, encryption and picture-manipulation. Plug-ins can be compared to regular include-files in Classic ASP, except they are only loaded when you initialize them. As soon as a plugin is instantiated, the plugin's-code is available in the ASP page's namespace.</p>

<p>It's possible to develop your own aspLite plug-ins, by adding asp files (including a VBScript class) in the appropriate folder: asplite/plugins/myplugin/myplugin.asp. The most important rule is not to response.write anything to a browser from within a plugin. That would make your plugin useless when working with asplForms.</p>

</div>

<div class="pagebreak"></div>
<a name="chapter7"></a>
<h2>Excursions</h2>

<p>To conclude this e-book, I want to discuss a couple of loose ends. Things I have been developing over the years, whether or not relying on JavaScript frameworks or relating to aspLite.</p>

<h3>Disconnected recordsets</h3>

<p>Disconnected recordsets are regular recordsets. But they have nothing to do with databases. They are full-blown datatables, stored in the computer's RAM. Disconnected recordsets are much smarter than VBScript dictionaries and arrays. They are designed to store massive volumes of data of any given data type.</p>

<p>aspLite comes with a helper-function that returns a disconnected recordset:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/recordset.inc")))%></pre>

<p>An example:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/drs.inc")))%></pre>

<p><strong>Live example:</strong></p>

<div class="asplForm alert alert-light" id="drs"></div>

<p>Disconnected recordsets are often used to convert large custom datasets for use in JavaScript frameworks, like could be the case in the example above.</p>

<p>The possibilities are endless. And there are a <a href="https://windows.epfl.ch/devguru/ado/recordset_fieldscollection.html" target="_blank">lot of different datatypes</a> you can store in a recordset. You can also sort, filter and delete records in a recordset. In the above example, the recordset is sorted by last name and first name (both ascending).</p>

<p>One use case for disconnected recordsets could be to list the 100 most recently modified files in a folder. By default, files are sorted alphabetically when browsing them. Adding the <code>file.DateLastModified</code> to a recordset, allows to sort files by modification date, instead of filename.</p>

<p>Also, sometimes you may want to create recordsets based on complex database queries, combined with other data from other data sources. The sky's the limit. Again, recordsets don't have anything to do with databases. You can use them to organize and export any data of any type.</p>

<h3>Datatables</h3>

<p><a href="https://datatables.net/" target="_blank">DataTables</a> sure is one of the most popular jQuery plug-ins. <i>"DataTables is a plug-in for the jQuery Javascript library. It is a highly flexible tool, built upon the foundations of progressive enhancement, that adds all of these advanced features to any HTML table." (source: <a href="https://datatables.net/" target="_blank">https://datatables.net/</a>)</i></p>

<p>These past few years, I have successfully used DataTables in various Classic ASP/VBScript applications, in various ways.</p>

<p>Long story short: all you need to get started with DataTables, is include both the CSS and JS libraries in your HTML header (head-section):</p>

<pre class="alert alert-light">&lt;!-- DataTables CSS &amp; JS --&gt;
&lt;link rel="stylesheet" href="https://cdn.datatables.net/2.0.3/css/dataTables.dataTables.css" /&gt;
&lt;script src="https://cdn.datatables.net/2.0.3/js/dataTables.js"&gt;&lt;/script&gt;</pre>

<p>From there on, things get very exciting very quickly:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/datatables.inc")))%></pre>

<p><strong>Live example:</strong></p>

<div class="asplForm alert alert-light" id="datatables"></div>

<p>This quick 'n dirty example demonstrates how a regular HTML table - filled with records from our person-table - miraculously turns into a searchable, sortable, pageable and striped DataTable. All it takes is the CDN-includes and:</p>

<pre class="alert alert-light">$(document).ready(function() {$('#myTable').DataTable()});</pre>

<p>Even though this technique is useful when you have few 100's of records, it will not work with large datasets as your ASP script would take too long to render the table-string. Therefore, DataTables also support AJAX-calls to load limited numbers of records based on searches/paging. You can preview a working sample of DataTables with AJAX for Classic ASP/VBScript <a href="https://samples.asplite.com/datatables/" target="_blank">here</a> and <a href="https://github.com/aspLite/DataTables" target="_blank">download its code</a> on Github.</p>

<h3>ChromeASP</h3>

<p>I developed ChromeASP in 2023, as I needed a solid PDF creator for a customer. ChromeASP is a free ASP/VBScript library that uses <strong>Headless Chrome</strong> (yes, Google's browser) to generate PDF and JPG files. The code and samples for ChromeASP can be <a href="https://github.com/PieterCooreman/ChromeASP" target="_blank">downloaded on Github</a>.</p>

<p>Live example: <a href="http://pdf.asplite.com/" target="_blank">http://pdf.asplite.com/</a></p>

<p>Classic ASP/VBScript developers never had an easy (and free) way to generate PDF and/or JPG/PNG files. This ASP script uses headless Chrome to get the job done. Chrome converts both websites (urls) and html-snippets to pdf and/or jpg (screenshots). As Chrome is creating the PDF files, there are absolutely no limits when it comes to HTML/CSS/JavaScript support. Chrome simply supports them ALL! Chrome even automatically fixes all sorts of errors and inconsistencies in your HTML. Even the most powerful and well known PHP PDF-libraries are looking at headless Chrome as their unbeatable successor. Check it out!</p> 

<h3>SheetJS</h3>

<p><a href="https://sheetjs.com/" target="_blank">SheetJS</a> is an Open Source JavaScript library that reads, edits and exports spreadsheets.</p>

<p>The aspLite <a href="https://demo.asplite.com" target="_blank">demo site</a> includes a working example. First, <strong>/default/asp/sampleform30.resx</strong> is loaded:</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("default/asp/sampleform30.resx")))%></pre>

<p>When the Submit-button is clicked, the following JavaScript is loaded by the browser:</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("default/html/sampleform30.resx")))%></pre>

<p>Apart from manipulating the button-class and -html, this JavaScript loads <strong>/default/asp/sampleform30_data</strong> to get data:</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("default/asp/sampleform30_data.resx")))%></pre>

<p>The function <code>aspl.json.dump(rs)</code> in <a href="default.asp?asplevent=sampleform30_data" target="_blank">/default/asp/sampleform30_data</a> generates the <strong>JSON-data</strong> needed by SheetJS to generate an XLSX-file in <code>function createFile(data)</code>. Long story short: very few lines of Classic ASP/VBScript and JavaScript code generate a pretty nice 5000-records (Unicode-proof) XLSX-file.</p>

<h3>Select2</h3>

<p><a href="https://select2.org/" target="_blank">Select2</a> presents itself as <i>"The jQuery replacement for select boxes"</i>. <i>Select2 gives you a customizable select box with support for searching, tagging, remote data sets, infinite scrolling, and many other highly used options.</i> (source: https://select2.org/) </p>

<p>Again, all it takes is to include the necessary CSS and JavaScript in head-section of your HTML page:</p>

<pre class="alert alert-light">
&lt;link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" /&gt;
&lt;script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"&gt;&lt;/script&gt;
</pre>

<p>Basic example:</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/select2basic.inc")))%></pre>

<p>This line: <code>$('#js-data-example-basic').select2();</code> turns a regular HTML selectbox in a searchable control. This is extremely useful for larger option-sets, like lists of countries, cities, etc.</p>

<p><strong>Live example:</strong></p>

<div class="asplForm alert alert-light" id="select2basic"></div>

<p>However, in some cases, it's impossible to pre-load all options in your selectbox. Imagine you want to be able to select a contact person out of a database with thousands of persons. You can't possibly preload such large option-sets in an HTML selectbox. Therefore, Select2 comes with an AJAX-solution:</p>

<p>Advanced example using <strong>AJAX</strong>:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/select2.inc")))%></pre>

<p>Apart from the select-control, the following JavaScript in <strong>ebook/select2js.inc</strong> is loaded:</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/select2js.inc")))%></pre>

<p>Select2 loads the following ASP script: <strong>ebook.asp?asplevent=select2data</strong>. It searches for matching results in our table "persons" (that we used in previous examples). Select2 sends the parameter "term" to <strong>ebook.asp?asplevent=select2data&amp;term=x</strong>. Our ASP script returns the JSON-object "results" at the very end of the script:</p>

<pre class="alert alert-light">aspl.dumpJson("{""results"": [" & suggestions & "]}")</pre>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/select2data.inc")))%></pre>

<p><strong>Live example (start typing 1 or 2 characters):</strong></p>

<div class="asplForm alert alert-light" id="select2"></div>

<h3>Bootstrap Modal</h3>

<p>A <a href="https://getbootstrap.com/docs/5.3/components/modal/" target="_blank">Bootstrap modal</a> is no doubt my favorite Bootstrap component. I can only think of advantages when using Bootstrap modals:</p>

<ul>
	<li>They keep a web application light and compact</li>
	<li>Modals work very well on smaller screens</li>
	<li>They come in different sizes (small, large, x-large, full-screen)</li>
	<li>Querystrings to modals are not exposed in the browser's address bar</li>	
	<li>They operate within their opener's document object model (DOM)</li>
	<li>They come with attractive slide-in animations</li>
	<li>They are a perfect fit for asplForm's AJAX capabilities</li>
</ul>

<p>The only limitation that I have faced in the past, is that you can only have <strong>one modal</strong> at a time. Another thing to keep in mind is that when you're having both large and small modals in your app, they could hide each other. The larger modals always hide the smaller ones. The only solution is to close (manually or by JavaScript) modals before loading new ones.</p>

<p>The necessary HTML for any Bootstrap modal is usually loaded (only once) together with the main HTML framework. The modal is nothing but an invisible div, somewhere in your body. For this occasion, I use a <strong>large modal</strong> (modal-lg). You can also have small (modal-sm) and x-large modals (modal-xl):</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/modal.inc")))%></pre>

<p>To <strong>open</strong> a modal, you can use::</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/modalbutton.inc")))%></pre>

<p><strong>Live example</strong></p>

<p><%=aspl.loadText("ebook/modalbutton.inc")%></p>

<p>The advantage of a <strong>static backdrop modal</strong> is that the user cannot close the modal by accident (by accidentally clicking it away). I always use static backdrop modals for interactive modals, where user-input is expected/required.</p>

<p>I slightly amended the HTML-code for this default modal. I added 3 id's for the 3 div's that matter:</p>

<ul>
	<li>id="modalTitle"</li>
	<li>id="modalBody"</li>
	<li>id="modalFooter"</li>
</ul>

<p>Let's load 3 asplForms for these 3 divs:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/modalload.inc")))%></pre>

<p><strong>Live example</strong></p>

<div class="asplForm alert alert-light" id="modalload"></div>

<p>Modals are by no means limited to short feedback messages. It's possible to load fully-featured interactive forms in Bootstrap modals. Let's revisit one of the first forms we used above:</p>

<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/modalloadfull.inc")))%></pre>

<p><strong>Live example</strong></p>

<div class="asplForm alert alert-light" id="modalloadfull"></div>

<p><strong>Closing</strong> a modal is as easy as:</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/modalfooter.inc")))%></pre>

<h3>Autocomplete</h3>

<p>I'd love to end this book where it all began, back in 2005. AJAX was invented at around that time and it became popular when Google started using it for its one and only input field on www.google.com. After all these years, nothing has changed. Google.com still offers just that one input field, serving you with auto-filled (autocomplete) suggestions and popular similar searches. That one and only autocompleting input field has made Google the strongest software company on the planet. Never change a winning team, isn't it.</p>

<p>This last excursion is about <a href="https://github.com/devbridge/jQuery-Autocomplete" target="_blank">Ajax Autocomplete for jQuery</a> and <a href="https://jqueryui.com/autocomplete/" target="_blank">jQuery UI's Autocomplete</a> and how to use them in Classic ASP/VBScript. Both jQuery libraries are in maintenance mode. They may even look a little outdated. They're no longer under active development. But they still get the job done!</p>

<h4>Ajax Autocomplete for jQuery by Devbridge</h4>

<p>As always, first things first: add the JavaScript library to the head-section:</p>
<pre class="alert alert-light">
&lt;!-- jQuery Autocomplete--&gt;	
&lt;script src="https://cdnjs.cloudflare.com/ajax/libs/jquery.devbridge-autocomplete/1.4.11/jquery.autocomplete.min.js"&gt;
</pre>

<p>Let's create a regular input field (type: "text") and load the necessary JavaScript:</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/autocomplete.inc")))%></pre>

<p>The following JavaScript in <strong>ebook/autocompletejs.inc</strong> is loaded:</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/autocompletejs.inc")))%></pre>

<p>The AJAX-call to <strong>ebook/autocompletedata.inc</strong> retrieves data (max 100 records) from our "person" table:</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/autocompletedata.inc")))%></pre>

<p><strong>Live example:</strong></p>
<div class="asplForm alert alert-light" id="autocomplete"></div>

<h4>jQueryUI's Autocomplete</h4>

<p><a href="https://jqueryui.com/autocomplete/" target="_blank">jQueryUI's Autocomplete</a> enables users to quickly find and select from a pre-populated list of values as they type, leveraging searching and filtering. </p>

<p>jQuery UI's Autocomplete is <i>nearly identical</i> to the implementation by Devbridge.</p>

<p>Include the necessary JavaScript/CSS in the head-section (below jQuery!):</p>

<pre class="alert alert-light">
&lt;!-- jQuery UI --&gt;
&lt;script src="https://code.jquery.com/ui/1.13.2/jquery-ui.min.js" 
integrity="sha256-lSjKY0/srUM9BE3dPm+c4fBo1dky2v27Gdjm2uoZaL0=" crossorigin="anonymous"&gt;&lt;/script&gt;
&lt;link rel="stylesheet" href="https://code.jquery.com/ui/1.13.2/themes/base/jquery-ui.css"&gt;
</pre>

<p>Let's create a regular input field (type: "text") and load the necessary JavaScript:</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/autocompletejQueryUI.inc")))%></pre>

<p>The following JavaScript in <strong>ebook/autocompletejs.inc</strong> is loaded:</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/autocompletejQueryUIjs.inc")))%></pre>

<p>The AJAX-call to <strong>ebook/autocompletejQueryUIdata.inc</strong> retrieves data from our "person" table:</p>
<pre class="alert alert-light"><%=pre(server.htmlEncode(aspl.loadText("ebook/autocompletejQueryUIdata.inc")))%></pre>

<p><strong>Live example:</strong></p>
<div class="asplForm alert alert-light" id="autocompletejQueryUI"></div>

<div class="mt-4 mb-4 p-4  text-bg-success">

<p><strong>Summary</strong></p>
<p>All in all it does not take a lot of Classic ASP code to integrate with very popular JavaScript frameworks and jQuery-plugins. And the final result is often sensational, especially compared to what we had available in the early years of Classic ASP. ASP may be outdated, but when using these JavaScript libraries, we can make our applications look great again.</p>

<p>The aspLite <a style="color:#FFF" href="https://demo.asplite.com" target="_blank">demo site</a> integrates some more JavaScript libraries: <a style="color:#FFF" href="https://jqueryui.com/datepicker/" target="_blank">jQuery UI Datepicker</a>,&nbsp;<a style="color:#FFF" href="https://stuk.github.io/jszip/" target="_blank">JSZip</a>,&nbsp;<a style="color:#FFF" href="https://github.com/MrRio/jsPDF" target="_blank">jsPDF</a>,&nbsp;<a style="color:#FFF" href="https://ckeditor.com/" target="_blank">CkEditor</a> and <a style="color:#FFF" href="https://codemirror.net/" target="_blank">CodeMirror</a>. Check them out!</p>

</div>


<div class="pagebreak"></div>
<a name="chapter8"></a>
<h2>Final notes</h2>

<p>Classic ASP/VBScript is fun. It always was. It's a capable, fast, server-friendly, light, easy-to-read and understand web development toolbox. In general, it still runs faster than both ASP.NET and PHP ever did on Windows Servers. MicroSoft should never have dropped it. It may not be as powerful, well thought of or structured as ASP.NET always was, Classic ASP had a large following and a reason to exist. It always was - and still is for some - a very popular scripting language.</p>

<p>But Classic ASP lost all its shimmer and shine over the years. The talented developers jumped off the MicroSoft boat a long time ago. I wasn't able to jump. I just had too much going on in Classic ASP back then. And all these years I kept on digging my development grave for way too long now. I need to make up my mind about my career as a web developer someday soon.</p>

<p>aspLite comes way too late in the game, and it's not much more than a collection of fresh ideas on how to use Classic ASP in 2024. In fact, it would have been a better idea to build aspLite on a true MVC architecture from the start. I may get that done any time soon.</p>

<p>Despite the many kudos I receive ever since I started building aspLite in 2020, I am probably the only developer who actually uses aspLite and asplForms. I'm still building amazing web apps with it, yes in 2024. My most recent project is <a href="https://setlistplanner.com" target="_blank">Setlistplanner.com</a>, a web app to organize songs into shareable setlists to ease band rehearsals and live gigs. Setlistplanner.com is built entirely on aspLite.</p>

<p>I recently even started offering good old Classic ASP applications as PWA's (Progressive Web Applications). There is no way you can tell the difference between a responsive PWA (using Bootstrap) and a native app on Windows PC's and/or Android devices. Apple is a different story though.</p>

<p>The cover of my book shows a recent photo - taken by my wife - of our 15y old daughter, left alone in the woods. Just like we are. Lost but hopeful to find our way out one day, sooner or later.</p>

<p>I hope you liked reading this book. I put all my heart and brains in it.</p>

<p>Pieter Cooreman, April 2024.</p>

<hr>

<div class="mt-4 mb-4 p-4  text-bg-secondary">
<img align="left" style="width:170px;margin-right:25px;margin-bottom:25px" class="img-thumbnail rounded" alt="Pieter Cooreman" src="ebook/pietercooreman.jpg" />
<p><strong>About the author</strong></p>
<p>Pieter Cooreman (born 1972) started his developer career in 1998. As a die-hard expert in Classic ASP/VBScript - the most popular web development technology between 1997 and 2003 - Pieter developed his way through basically any type of web application. Between 2007 and 2014, Pieter developed and supported QuickerSite webCMS, a free ASP/VBScript CMS loved and still used by hundreds. At its peak, worldwide more than 6000 QuickerSites were hosted on a variety of hosting solutions for a wide variety of customers and industries.</p>
<p>As Classic ASP is still around, Pieter always loves to share ideas through ASP code and snippets. In 2020 the pieces of a complex puzzle started to fall in place with aspLite, an AJAX-first developer framework for Classic ASP/VBScript developers. Writing this book felt like a roller coaster and sure led to new ideas. Classic ASP/VBScript may be dead, but will be hostable for at least another decade on Windows Servers. So we can better have fun with it. Because fun is key!</p>
<p>Besides being a self-employed developer, Pieter is a caring father of three and a loving husband, living near Leuven, Belgium. He's a web developer during winter and a highly demanded live music-performer during summer, having regular appearances on national radio and tv. Pieter is living his dreams and passions. And sharing them - Pieter believes - brings true happiness.</p>
</div>

</div>


<%=aspl.loadText("ebook/modal.inc")%>
</body>
</html>