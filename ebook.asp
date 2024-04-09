<!-- #include file="aspLite/aspLite.asp"-->
<%
dim asplEvent : asplEvent=aspl.getRequest("asplEvent")

if not aspl.isEmpty(asplEvent) then
	
	'dynamically execute the scriptname in asplEvent
	aspl("ebook/" & asplEvent & ".inc")

end if
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
	
	<!-- Bootstrap CSS & JS -->
	<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-QWTKZyjpPEjISv5WaRU9OFeRpok6YctnYmDr5pNlyT2bRjXh0JMhjY6hW+ALEwIH" crossorigin="anonymous">

	<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-YvpcrYf0tY3lHB60NNkmXc5s9fDVZLESaAA55NDzOxhy9GkcIdslK1eN7N6jIeHz" crossorigin="anonymous"></script>
	
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

<div class="mb-4 lead">

<h1>aspLite: a future for ASP/VBScript</h1>

<p>This book is about Classic ASP and VBScript and how they changed my life in 1999. And how MicroSoft took it back in 2002.</p>

<p>If you want to discuss the various topics covered in this book, do not hesitate to start a discussion on <a href="https://github.com/aspLite" target="_blank">https://github.com/aspLite</a></p>

<p class="text-end">Pieter Cooreman, April 2024</p>

</div>

<h2>Preface</h2>

<p>Back in 2002, thousands - difficult to say how many exactly - of ASP/VBScript developers were left alone in the woods by a very small team of MS developers who choose to build ASP.NET and throw ASP and VBScript in the bin. ASP got cancelled, from one day to the other. That's how it felt back then, and that's what really happened looking back on it today.</p>

<p>I sometimes try to picture it like this: In 2002, Microsoft had just had a great couple of years. Windows 95, 98, 98SE, 2000 Professional, NT Server, 2000 Advanced Server. Not to mention Windows Me and the various XP editions. MicroSoft also had Internet Explorer, Office, SQL Server and many other products. These were all great products that were loved by basically any gamer, student, school, non-for-profit and  all businesses. Between 1997 and 2002, MicroSoft was the leading software company worldwide, without any doubt. They sure could afford themselves to make mistakes and make some wrong decisions. And so they did. That's how Classic ASP/VBScript ended up in the bin, without much ado. Just like that.</p>

<div class="mt-4 mb-4 p-4 text-bg-dark lead">

<p>Ever since 2002, Classic ASP/VBScript developers live in the Dark Middle Ages. We're being ignored, laughed at, made fun of and expelled from MicroSoft's Promised Land.NET. We're loose ends, lone wolfs. We live in an endless and sad diASPora, cut off from the Mothership. This hurts. This has to stop. Rather sooner than later.</p>

</div>

<p>ASP/VBScript were very popular web development and scripting technologies between 1997 and 2002. They were MicroSoft's first answer to another rapidly growing web technology and competitor: PHP and Apache servers. We all know how that turned out. However, most businesses preferred ASP over PHP back then. ASP was running faster, integrated well with other MicroSoft products (Active Directory), had a rapidly growing community and supported (multi-threaded) COM-extensions. But above all, ASP/VBScript was very easy to learn, even for people like me, without a degree in computer science whatsoever. VBScript was a visual, case insensitive and lousy typed scripting language after all.</p>

<p>Even though ASP was a huge success, MicroSoft developers decided to pull the plug and amuse themselves with... how many ... well what the heck, 25 versions of ASP.NET. ASP.NET never was a popular web development technology. In it's early years it was left far behind by PHP frameworks. Today full stack JavaScript frameworks have taken over. ASP.NET is and never was nowhere near a popular or frequently used web development framework. ASP/VBScript was.</p>

<p>PHP developers had more luck. Between 2000 and 2010 (about the same period of the rise and fall of ASP.NET and Windows Servers) a handful of PHP developer frameworks gained popularity amongst web developers (Zend, CakePHP, Symfony, Laravel, ...). They were there to stay. Everybody felt that. Unlike ASP.NET, these PHP frameworks were <i>built</i> with PHP, they did not want to <i>replace</i> it.</p>

<p>ASP/VBScript developers also needed a developer framework back in 2002. They did not need ASP.NET. A developer framework for ASP/VBScript should have taken care of various shortcomings in ASP/VBScript, workarounds for known bugs or issues, facilitate code behind, become event-driven, bring a solution for the spaghetti-coding, improve coding habits, increase scalability and security. Last but not least, ASP/VBScript needed a framework in order not to reinvent the wheel each time a new project was to be developed. But that framework never happened.</p>

<p>Modern and popular web development frameworks like Django and Vue are also facilitating the so-called spaghetti-way of coding. It does not prevent those frameworks to gain a large following amongst both beginning and more experienced developers. Same was 100% true for Classic ASP back then.</p>

<p>By the time ASP.NET was stable (not before 2.0 in 2005), most talented developers had given up on .NET and went developing plugins for WordPress, Joomla or Drupal, or they developed one or the other social network. Using PHP. For me it was very simple: By then, I had 100s of customers who were using my Classic ASP applications. I built an entire hosting business around ASP/VBScript somewhere between 2000 and 2005. I did not want to refactor 10000s of lines of code just because MicroSoft wanted me to. So I got stuck with Classic ASP. And even today in 2024, I'm still hosting 100s of web applications solely relying on ASP/VBScript.</p>

<p><strong>Classic ASP was considered sloppy, buggy and weak. A technology for lesser Gods.</strong></p>

<p>At least that's what MicroSoft developers wanted us to believe back then. And I sure identify as a lesser developer God. I'm lucky enough to have other talents.</p>

<p>We did have some individuals who did a great job with some very powerful scripts though. I remember a pure ASP/VBScript upload class by Lewis Moten that I used a lot in <a class="link" target="_blank" href="https://quickersite.com">QuickerSite</a> and other projects. It got fine-tuned and Unicode-proof by other developers along the years. Others developed very useful encryption classes for Classic ASP (Sha256, MD5). I remember a very useful Captcha image generator-class. I once developed a class to facilitate Google Maps in ASP applications. I also developed some classes that I always reuse when starting new ASP applications. They relate to security, login, sessions, cookies, database connectivity and CRUD-statements, etc. </p>

<p>But the Classic ASP community also saw some great attempts to develop true frameworks. Michal Gabrukiewicz did a great job with <a class="link" target="_blank" href="http://www.asp-AJAXed.org/">asp-AJAXed</a> but sadly passed away in 2009. More about Michal later. He played an important role when developing aspLite, even though he passed away 10 years before I started developing it. <a href="https://www.codeproject.com/Articles/7922/Classic-ASP-Framework-2-0-Make-your-Classic-ASP-co" target="_blank">CLASP</a> by Christian Caldereon was another great .NET-alike framework for Classic ASP. All in all, developer frameworks for ASP/VBScript developers did not last long and were often developed by a single developer. Apparently it was very hard to build a community around Classic ASP back then. I still blame MicroSoft. MicroSoft has given its own ASP/VBScript tandem very bad press back in those days. Nobody dared to stick his neck out by building a durable ASP/VBScript framework back in 2003-10. I still regret that.</p>

<p><strong>No painless nor hassle-free upgrade-path to ASP.NET</strong></p>

<p>Back in those days, ASP/VBScript developers did not have a painless nor hassle-free upgrade path to ASP.NET. VBScript was not supported by ASP.NET. Oh yes, for small web applications, one or two ASP pages, it was doable. MicroSoft even provided some automated conversion tools. But they were very limited. In 2003, by the time ASP.NET 1.1 fixed some very annoying bugs in 1.0, I was dealing with 3 extremely large ASP/VBScript code bases. Tons of includes files, classes, functions and routines. It was impossible to refactor them to ASP.NET without spending at least a year. And for what reason? ASP.NET did not offer much extra compared to ASP/VBScript at that time. In 2002, ASP.NET was often considered <i>too little too late</i> by seasoned MicroSoft developers. But what the heck, I didn't have the time for all that. I was building a business around my developing, selling and communication skills. And I really didn't need ASP.NET for that.</p>

<p>As soon as a technology gets introduced and gains a large user base - as was the case for ASP/VBScript between 1997 and 2002 - it's impossible to stop it. People will always be prepared to live with its limitations, work their way around them or learn to live with them. That's exactly what I did back then, and I'm still doing now with aspLite. And I'm not alone. It's a human thing. And we're all humans after all. MicroSoft misjudged that.</p>

<p>We're 25 years later now. There are still loads of Classic ASP/VBScript applications out there, most of them serving dynamic websites for more than 20 years now. Still MicroSoft refuses to clearly communicate about the EOL policy of Classic ASP. There IS no EOL policy for ASP/VBScript. So after all these years, we - ASP/VBScript developers - are STILL left in the woods. Alone. Without even the slightest clue on when exactly MicroSoft will pull the plug.</p>

<p>I learned to live with that even though I still feel sad about it. In 2020 I decided to develop a new framework for ASP/VBScript developers: <a class="link" href="https://aspLite.com" target="_blank">aspLite</a>. I will be working on aspLite for the rest of my life. I somehow love this technology. And it just won't end. Classic ASP is fun! And fun is key.</p>

<p><strong>Some kudos for MicroSoft though</strong></p>

<p>Sure, it's not all bad news either. MicroSoft still offers Classic ASP through ISAPI in IIS and it made sure that all our ASP applications just kept on running smoothly like they did back in 2000. They even run better now, thanks to today's faster hardware. MicroSoft did not mess up Classic ASP. They did not add things to it, but they also did not screw things up. That's a plus.</p>

<p>We also saw some interesting attempts to provide an IDE for Classic ASP developers. We had two flavors of "Webmatrix", an all-in-one IDE for both MicroSoft and Open Source technologies. I liked both versions of Webmatrix a lot. But they both suddenly disappeared at some point in time.</p>

<p>So far, all Visual Studio editions supported intellisense and code-completion for Classic ASP. And IIS Express supports Classic ASP to the full 100%. Today, even the Windows 10 and 11 Home editions come with a full version of IIS. In a way, nowadays it's much easier to start developing in ASP/VBScript than it was back then. You needed Windows 2000 Professional or a Server back in 2000. Today you only need to know how to get Classic ASP up and running. Unfortunately, very few Windows users know how, and they couldn't care less. This illustrates how things turned out for MicroSoft. In 2000, companies spent quite some money on Windows 2000 Pro licenses. Today, MicroSoft ships its entire development framework and its dependencies for free. But even that never turned the tide.</p>

<p>Another pleasant evolution is that - over the years - MicroSoft embraced Open Source technology. At some point in time, it was easier to install both WordPress, Joomla and Drupal on any Windows host than it was to get ASP up and running. The wonderful Web Platform Installer however, was retired in 2022 - again - for inexplicable reasons.</p>

<div class="mt-4 mb-4 p-4 text-bg-warning lead">

<h4>Classic ASP end of life? Fake news.</h4>

<p>Some sites report that Classic ASP is "end of life". The funny thing about this is, these are articles written between 5 and 10 years ago. Even <a href="https://en.wikipedia.org/wiki/Active_Server_Pages" target="_blank">Wikipedia</a> reports that "<em>ASP was supported until 14 January 2020 on Windows 7.</em>"</p>

<p>This is fake news. There is no official EOL policy for Classic ASP. As from IIS7, Classic ASP is implemented as an ISAPI filter in IIS, configured to kick-in ASP.DLL as soon as an *.asp file is requested. It's therefore more accurate to say: Classic ASP will be supported as long as IIS supports ISAPI-filtering. Or even better: the end-of-life of Classic ASP is inseparable from the end-of-life of IIS itself. That will be the day MicroSoft pulls the plug on its Server-products.</p>

<p>Same story for VBScript. VBScript is as dead as ASP, but it's actually a (default) part of <a href="https://en.wikipedia.org/wiki/Windows_Script_Host" target="_blank">Windows Script Host</a> (wscript.exe). Again, there is not a single reason to believe that VBScript will no longer be supported in WSH in the future. WSH is part of the Microsoft Windows Operating System ever since ... Windows 95.</p>

<p>Both Classic ASP and VBScript are <em>underlying&nbsp;</em>technologies. ASP is a toolset for web developers. VBScript is a visual programming language. Other services and software depend on them (both MicroSoft and non-MicroSoft).&nbsp;They're not to be confused with end-user products that need security-fixing, updates, support, legal follow-up,&nbsp;etc. Classic ASP nor VBScript are even included&nbsp;in the "<a href="https://learn.microsoft.com/en-us/lifecycle/products/" target="_blank">Microsoft Products and Services</a>" lifecycle database. IIS is <a href="https://learn.microsoft.com/en-us/lifecycle/products/internet-information-services-iis" target="_blank">included</a> though. It follows the&nbsp;Component Lifecycle Policy, meaning that it's supported as long as the product where it's a component of is supported. That is the Windows Server family.</p>

<p>Therefore, I truly believe that Classic ASP and VBScript will be available in Windows OS for at least another 10-15 years, probably longer. Nobody knows what happens next. This is also true for any other technology you'd use today. So there is not a single reason to not consider Classic ASP/VBScript to develop web applications. I can tell. I'm still doing so. And I love it. As long as someone is coding ASP/VBScript, it's alive. I must admit however that I'm more and more feeling lonely. It appears that most web developers don't even know what Classic ASP exactly is/was. I'm trying to turn the tide. But that will never work on my own...</p>

</div>
<div class="pagebreak"></div>
<h2>Story behind aspLite</h2>

<p>The story behind aspLite tells the story of my career and the way I dived into web development about 24 years ago.</p><p><strong>The early days</strong></p><p>24 years ago - in 2000 - I was 28 years old. I was young and eager to learn to code. I had no degree in computer science whatsoever, but I picked up a lot from colleagues very quickly. Developing web applications quickly became an obsession. I developed all sorts of applications &nbsp;- both as a professional web developer, and for a hobby. Back in those days, I happened to work for a company that specialized in developing COM components for e-commerce websites. I was not part of the RnD team actually building these components. I was part of the support-team, implementing them for customers. We used Classic ASP and VBScript. What else did you think?</p><p><strong>Visual Basic Scripting - VBScript&nbsp;</strong></p>

<p>Rather than use the COM components made by that RnD team, I quickly realized that it was actually much easier (and much quicker) to develop custom classes in VBScript to fully meet the customer's requirements. Actually, using those COM components slowed down our applications and the development cycles. So in the end, we didn't use them. It somehow meant the end of that company. And I was in it for something... well, a lot actually. When I left that company in 2002, I took it's biggest customer with me and started my own company. Shame on me. But hey... that's life. I was a 30y old entrepreneur after all. I simply had to do it.</p>

<p>VBScript was the first programming language I learned to use. Full stack developers will claim that VBScript is useless and cannot be called a serious programming language. I disagree. VBScript is <strong>visual </strong>(easy to read/write, case-insensitive coding, no nested curly braces {{{{{}}}}} - I mean...), <strong>basic</strong> (easy to understand, no complex statements) and <strong>scripted</strong> (just-in-time, no compilation). These 3 properties made a <em><strong>huge success</strong></em> of VBScript back in 1997-2002. VBScript can also be used together with ActiveX Data Objects (ADO) - a high-level, easy-to-use interface to OLE databases (Access, SQL Server, Oracle, etc). ADO is what made VBScript a success. And it still does.</p>

<p><strong>QuickerSite</strong></p>

<p>After having coded my way through basically any type of web application somewhere between 2000 and 2007, I decided to come up with a CMS (Content Management System) that combined all my best scripts and coding habits that had passed the test of time so far: <a href="https://www.quickersite.com/" target="_blank">QuickerSite</a>. QuickerSite was a success, especially in it's early years. Only one year after it's initial release in 2007, it was translated into 11 languages, including Danish, Hebrew, Italian, Turkish and Swedish. In 2008 QuickerSite was used by about 1000 users worldwide, who created at around 6000 QuickerSites in total. It was as if a lot of ASP coders all over the world had been waiting for a CMS developed in pure ASP/VBScript.&nbsp;</p>

<p>Between 2007 and 2014 I built a hosting business around QuickerSite. At its peak, I hosted 1200 QuickerSites on a single dedicated <a href="https://www.quickersite.com/r/default.asp?iId=EDJEGEIH&iPostID=EHILMLH" target="_blank">Dell server</a> with very basic specs (3GB RAM, a slow 120 GB SATA disk and a single CPU). But it worked. It rocked. And all this time, I was a one-man-band. Nobody else but me, myself and I was dealing with everything related to my business: selling, developing, designing, hosting, invoicing, mailing. And everything else related to QuickerSite. So yes, I sure was a bit of a lone wolf back then. And as much as Microsoft forced me to switch over to ASP.NET, I did the exact opposite and stayed with Classic ASP. That's me.</p>

<p>Developing, selling, hosting and supporting QuickerSite for a wide variety of customers (and hosting conditions) learned me even more about ASP/VBScript, its caveats AND its strengths. Developing QuickerSite was no doubt the best time of my life. I still host many QuickerSites today on a Windows 2019 Server in the cloud. And I made some life-time friends.</p>

<p><strong>The 2010 paradigm shift</strong></p>

<p>By the time QuickerSite had grown mature - somewhere in 2010, there were quite a few things going on: HTML5, CSS3 and JavaScript frameworks got adopted by the WWW, mobile devices (phones and tablets) were rapidly gaining in popularity, social media took over our lives, and last but not least - by then - a large group of developers adopted open source solutions and frameworks developed in PHP/MySQL. Nowadays, JavaScript/CSS frameworks like Bootstrap (Twitter), React (Facebook), Angular (Google) and Node.js are completely dominating the web development business. While Microsoft was trying hard to keep up with these ever changing circumstances - by releasing dozens of ASP.NET versions and editions - Classic ASP developers were slowly and silently being ignored and left alone in the woods. Many of them are retired by now, or they are no longer actively developing (new) Classic ASP applications. I regret that a lot, because all this time - and still today - ASP/VBScript has been a perfect fit for these JavaScript frameworks, for instance by providing very straight-forward database access (ADO), dealing with binary files (uploading/streaming) or by simply delivering very useful AJAX, XML and JSON integrations. All that is where aspLite is about.</p>


<div class="mt-4 mb-4 p-4  text-bg-primary lead">
<p><strong>Trip down memory lane</strong></p>

<p>When Jesse James Garret "invented" AJAX in 2005 - and Google made a success of it - some Classic ASP/VBScript developers must have thought by themselves: what the heck, we did AJAX back in 1999 already! We used something named <strong>Remote Scripting</strong>. It was AJAX avant-la-lettre. Nothing more, nothing less. We loaded a Java Applet (it took about 3 seconds to load) and next we were able to provide AJAX functionality in pure Classic ASP/VBScript. Back in 1999. This technique worked fine on both Internet Explorer and Netscape, the only two browsers that mattered back then.</p>
</div>

<p><strong>PHP and ASP.NET libraries in Classic ASP</strong></p>

<p>As Classic ASP is a dead-end street anyway, it may be a good idea to do some neighborhood shopping. Why not use some PHP or ASP.NET libraries in Classic ASP? I'm doing that for many years already. I use .NET's web.config files to configure url rewriting (http-&gt;https), set custom error handling (404 catch) and set default documents (default.asp). I also a developed a single VB.NET page that takes care of (unobtrusive) server-side image-resizing/cropping and recoloring. I added that aspx-page to the jpg-plugin in aspLite (jpeg.aspx). The idea of using resx extensions for HTML files is another (ab)use of .NET. All in all it's not much, but it is something.</p>

<p>In the past I have also successfully implemented PHP's built-in zipper (ZipArchive()) and a plugin named mPDF to create zip -and pdf files on the server from within a Classic ASP application. Worked like a charm. As these libraries reside on the server, they can safely be implemented with internal ServerHTTPRequests from a Classic ASP environment.</p>

<p><strong>Concerns about the future of Classic ASP</strong></p>

<p>Some developers may ask themselves: "<em>For how long are we going to be able to host Classic ASP applications on Windows Servers</em>"? Windows 2022 Server fully supports all Classic ASP/VBScript components, better than ever before (also thanks to much better/faster hardware). As soon as Microsoft releases his next Server edition, I'll be amongst the first to deeply test-drive Classic ASP on it.</p>

<p><strong>Suggestion: turn ASP/VBScript into Open Source</strong></p>

<p>The best way forward for Classic ASP is to turn it into an open source project, pretty much like Microsoft did with <a href="https://github.com/dotnet/aspnetcore" target="_blank">ASP.NET Core</a>. Given today's success of scripted languages like JavaScript, Ruby and PHP, VBScript still has the potential to reach a different kind of developer. Some developers (like me) will NEVER be able to read through dozens of nested curly braces. It's just visually too poor for me. But that does not mean we are unable to write poetry in ASP/VBScript. I'll keep on trying anyway.</p>

<div class="mt-4 mb-4 p-4  text-bg-primary lead">

<h4>What and who I'm doing this for...</h4>

<p>I often receive messages like the ones below. It's striking how many of us Classic ASP developers share the same feeling.&nbsp;Left alone in the woods with our ASP-scripts and functions that passed the test of time (some of them are working flawlessly for 25 years now, on all Windows Servers).</p>

<p><i>"I am moved with your back stories with the development of aspLite. I am an ASP web developer since 2004 and still using asp Classic as for my freelance projects. I can relate on the struggle on having shared hosting upto know and I would love to learn more about AWS EC2 instance and own a hosting server like what you did. I'm happy to know that there are still believing in the power of Classic ASP."</i></p>

<p><i>"I like your work! Contrary to what the muppets say, Classic ASP is not dead. There is nothing that one cannot do with ASP. In fact I have written desktop applications in ASP... cannot do that with PHP. You may recall SOOP? It had a huge following and a lot of devs for a lot of useful plugins. Not sure what happened with that even though I was one their plugin devs. Perhaps it was the cost of web hosting vs Apache servers and WordPress. Good luck with your CMS."</i></p>

<p><i>"I've just found your web site for the first time today and I just wanted to say thank you for doing it. I've been using Classic ASP since the late 90s and still use it today for projects even though I've been jeered and sneered at for doing so. It's great to see there are other people still out there using it for real and still flying the flag :0)"</i></p>

<p><i>"I found your page by chance, because I was looking for something in ASP that would work with the "DataTables | Table plugin for jQuery"... congratulations for maintaining this site. Hope you are still working with ASP...thanks."</i></p>

<p>Out of all emails I have ever received, I will always remember this one. It's about QuickerSite, but it also applies to Classic ASP/VBScript... in a way:</p>

<p><i>"Pieter, many years ago I came across Quickersite and I payed I think for something (I hope I was a payer and not just a downloader) and I was overwhelmed with how amazing it all was. I thought I'd found the holy grail for me being a provider of web services... but my history has followed a similar trajectory to QS and I'm now wandering the plains with a lot of other buffaloes wondering what to do next. I read your posts occasionally as they are, and for a weird reason - it gives me hope. Hope because what you had going here was SO ON THE MARK. But the world rotated a little in another direction and twitter/FB / Instablah and all the other mindless platforms took the collective consciousness away from individual effort and relevant opinion, such that QS no longer had a home. ... Why? Who knows, the world is just different now and windows of opportunity open and close for the most inexplicable of reasons. This could of course be used as an equal answer to: "Why does humanity exist"? A: "For inexplicable reasons". So; no more Concorde, no more digital watches, no more moon landings and no more QS. It is in good company and I loved it."</i></p>

</div>

<h4>Some final notes before diving into the code</h4>
<!--
<p><strong>Access databases?</strong></p>

<p>A lot has been said about using Access databases in web applications (whether or not developed in Classic ASP). Bottom line, most experts say: <em>don't</em>! However, I say: Access is the easiest, fastest, most server-friendly (uses <strong>no RAM</strong>!) and an extremely reliable database to use with Classic ASP. Period. I'm using them for 20 years now. Never ran into a single problem. And today - with these fast SSD drives and powerful CPU's - you're not going to crash an ASP page that's using an Access database. I have gone through all the available stress-tests over the years. I was even getting frustrated at some point, as I have invested a lot of time and money in migrating to SQL Server in the past. All in vain. Waste of time and money. I went back to Access for all my hosted websites in 2018. No regrets.</p>

<p>That said, there are some guidelines and things to keep in mind when using Access in Classic ASP:<ul><li><u>Do not store</u> <strong>binary files</strong> (images, pdf, documents) in Access databases. Store them as regular files. If you're concerned about security, give your files a secure (not downloadable) extension (like .resx) and stream them to users with <em>adodb.stream</em>. aspLite facilitates this with its <em>dumpBinary </em>(see <a href="api.html">API</a>) function.</li><li><u>Do not store</u> <strong>visitor data</strong> (logfiles) in a database. Again, use the file system. There is no point in storing bulk data that you're probably never going to look into afterwards anyway. Also, visitor data are typically stored in the IIS logfiles already. No need to duplicate them.</li><li><u>Do not store</u> <strong>data-backups</strong> in the database itself. Some developers log each and every change to database-records in yet another (copy-)table in order to keep track of changes. In some cases that could be useful. But when using Access, it's very easy to keep a couple of backups to revert to or look into in case data may have been lost or corrupted. Again. Do not store bulk data in an Access database if you're not gonna need it. Make backups... I have a customer on a 250MB Access database. Manually making a copy of that file takes... <em>half a second</em> on a fast SSD drive. So there is no reason to not make Access backups every day, or even every couple of hours.&nbsp;</li><li>Summarizing: <em><strong>only store text and numeric data</strong></em> in Access databases that you're actually going to <em><strong>use</strong> </em>in your application: read, update, delete and search for.</li><li>Make <strong>backups</strong> of your Access mdb files. Do it. Every day, or even every couple of hours for business critical applications. You really don't want to lose data.</li><li>Have a look at the <strong>database-plugin</strong>&nbsp;: /aspLite/plugins/database/database.asp. Ideally, an ASP page opens a connection to a database only ONCE through its lifespan. Opening (and closing) database connections are probably the most time-consuming operations in ASP pages. Doing this only once drastically speeds up your ASP pages. </li><li>There is a limit of <strong>255 concurrent users</strong> for Access databases. However, when using Access in a web application, the "user" (IUSR in most cases) is only connected for a few milliseconds. You wil not often face situations where 255 visitors simultaneously open an ASP page <em>(that is: on the very same millisecond)</em>. Unless you're Google or Facebook... I have never faced that situation in 20 years time.</li><li>There is a limit of <strong>2GB </strong>for Access databases. That's a lot of text and numeric data. Do not let Access databases grow <em>even close</em> to 2GB. I would personally decide to upsize to SQL Server as soon as an Access mdb file grows bigger than 500MB. But that's just me, given the hardware I use and trust (EC2 instance on AWS). I currently do not host a single Access database bigger than 250MB.</li></ul></p>
-->

<p><strong>IDE</strong></p>

<p>Classic ASP developers have been lacking a dedicated IDE (Integrated Development Environment) for some time now. I therefore prefer to use&nbsp;<a href="https://notepad-plus-plus.org/" target="_blank">Notepad++</a>. It's free, lightning fast, reliable and it's very easy to use. Notepad++ also comes with some basic code-completion functions. Notepad++ can also be freely installed and used on any Windows Server. I often use it to search for specific texts and strings in over 1 million files. No problem at all. Others may want to use <a href="https://code.visualstudio.com/" target="_blank">Visual Studio Code</a> instead.</p>

<p><strong>Where to host ASP sites today?</strong></p>

<p>A final note on hosting. I have personally never used or liked <em>shared </em>ASP hosting solutions of any type. From day one, I use my own server. In 2004 I bought my own. In 2010 I migrated to the cloud. Today (2024), I use an&nbsp;<a href="https://aws.amazon.com/ec2/" target="_blank">AWS EC2 instance</a>. Very satisfying so far. In my opinion, if you're into ASP development, you're better off managing a Windows Server <em>yourself</em>, rather than rely on shared hosting. There are very few people left who can assist you with ASP hosting issues. We're on our own. But that shouldn't be a problem as Windows Servers are very easy to setup, deploy and maintain. As an ASP developer you need to know the basics of setting up backups, firewall, IIS, mailserver and security on Windows Servers after all. On top of that, when you dive into ASP developing, you may have to install specific COM software or you may want to prefer a specific setup to facilitate code-reuse. This requires full access to a server. Both Microsoft Azure, Google Cloud and Amazon Web Services offer free-tier solutions that allow to test-drive a basic setup for a year. It's really worthwhile looking into these solutions.</p>

<p><strong>Should I learn Classic ASP/VBScript in 2024?</strong></p>

<p>If you're young and you're willing or able to dive into new technologies, do not learn Classic ASP or VBScript. If you want to become a web developer (by extension a mobile app developer) - today or in the future - learn Python and/or JavaScript. However, if you are already using Classic ASP (and plan to keep on doing so), it's always a good idea to learn new things. Even after all these years, there's always something new to learn about ASP. aspLite is living proof of that.</p>

<p>That said, learn ASP actually means: learn VBScript! There is not much to learn about Classic ASP after all, except for 5 objects with each a handful of properties and methods: Application, Session, Response, Request and Server.  VBScript however is the preferred (and default) programming language for most ASP developers. Even though VBScript is no longer being developed by Microsoft, it is still used a lot, not only by web developers, but also to automate tasks on Windows servers.</p>

<p>Out of many online resources you find when searching for "learn Classic ASP", I personally liked '<a href="https://asplite.com/assets/files/A%20Practical%20Guide%20to%20ASP%203.0.pdf" target="_blank">A Practical Guide to Microsoft Active Server Pages 3.0</a>' by Manas Tungare a lot. He was so kind to allow me to redistribute this guide on this website. These 70 well written pages are about everything you need to get started using Classic ASP and VBScript.</p>

Furthermore, W3Schools.com (that website is developed in Classic ASP) is probably the best online resource to start as a Classic ASP/VBScript web developer. Before you learn Classic ASP, you must learn HTML and CSS (and grab some basic JavaScript as well). Make sure to also check their very complete VBScript and ADO reference. ADO can be used to access databases from Classic ASP pages. </p>

<p><strong>About this book</strong></p>

<p>This book ain't really a book. It's an ASP script, using aspLite as its preferred framework. It appears that the easiest way to write a book about aspLite is using it while I'm trying to explain it.</p>

<div class="pagebreak"></div>

<h2>Where does aspLite fit in?</h2>

<div class="mt-4 mb-4 p-4  text-bg-success lead">

<p class="lead">24 years ago, Classic ASP was about Request() and Response(). Ins and outs. Simply put.</p>
<p class="lead">Today, any server-side web framework is still about just that.</p>
<p class="lead">For everything else you can - in 2024 - rely on JavaScript frameworks. No more need for spaghetti-ASP or HTML server-controls.</p>

</div>

<p>aspLite is a lightweight framework for ASP/VBScript developers. It can help to develop better ASP/VBScript applications. aspLite does not reinvent, replace or rewrite ASP components. It only tries to come up with a new way of using some of them.</p>

<p>ASP/VBScript (also known as "Classic ASP") is available for more than 20 years now and I believe it still makes sense to share this approach. After all, Classic ASP has proven to be reliable, fast and secure. </p>

<p>But that's not all. Today's web applications use JavaScript frameworks (Angular, Vue, React, jQuery, etc). Often only data (JSON) is transmitted back and forth to the server. That's why a server-side technology better stays small, fast and simple. Hence the name "aspLite".</p>

<p>An example. The <a href="https://demo.asplite.com" target="_blank">aspLite demo</a> site ships with a (fully functional) Classic ASP implementation of DataTables. This wonderful (free-to-use) widget has all it takes to offer ... datatables, including client-side sorting, searching and paging. There is only very little ASP code involved. Quite fascinating!</p>

<p>The vast majority of these JavaScript plugins are free to use, backed by 100's of contributors and they work in all (current) browsers, on all devices. How amazing is that?! Ever since the adoption of HTML5, CSS3 and ECMAScript 5 (somewhere between 2009 and 2012), client-side JS/CSS/HTML frameworks have become very popular. Much more popular than their server-sided predecessors (ASP(.NET), PHP, etc). Today, the most starred repositories on GitHub are all about (learning) JavaScript.</p>

<p>The <a href="https://demo.asplite.com" target="_blank">aspLite demo</a> puts the following front-end HTML/CSS/JS frameworks/plugins at work: Bootstrap, jQuery, jQuery UI Datepicker, JSZip, SheetJS, jsPDF, CkEditor, CodeMirror and DataTables. </p>

<div class="alert alert-success"><p class="lead">aspLite does not rely on 3rd party COM components. It only relies on the built-in VBScript components. Therefore aspLite works on basically each and every shared hosting solution out there, but also on each and every Windows host with ASP installed.</p></div>

<h4>Install aspLite as a virtual directory for code-reuse reasons?</h4>

<p>If you're lucky enough to manage a Windows Server yourself, you may want to install aspLite as a <i>virtual directory</i> in your IIS website. This way you can use one aspLite codebase on an unlimited number of websites on your server. If you use a network location, you can even share one aspLite codebase amongst multiple servers.</p>

<p>When setup this way, the first line in your ASP page would read:<br>
<code>&lt;!-- #include virtual="aspLite/aspLite.asp"--&gt;</code> (in case you name the virtual directory "aspLite")</p>


<h4>Getting started</h4>

<p>In its most low-level mode, aspLite is nothing more (or less) than a library of ASP/VBScript classes, functions and subroutines. They can be found in /aspLite/aspLite.asp. I will go through all of them later on in this article.</p>

<p>This page - <strong>ebook.asp</strong> - you're looking at, includes the aspLite framework in it's very first line:<br>
<code>&lt;!-- #include file="aspLite/aspLite.asp"--&gt;</code></p>

<div class="alert alert-danger lead">It's extremely important that this so called include-file is ALWAYS the very first line in your ASP script.</div>

<p>Let's have a closer look at <strong>aspLite/aspLite.asp</strong>. The  first line of that ASP page reads: <strong>&lt;@LANGUAGE="VBScript" CODEPAGE="65001"%&gt;</strong>. That's probably the most important line in the entire aspLite code base. What is it doing?
<ul>
	<li>Set <strong>VBScript</strong> as the default server-side scripting language</li>
	<li>Ensure 100% compatibility with the <strong>UTF-8 character</strong> (Unicode) set, allowing you to use ANY language in your application and avoid very annoying character conversions and/or encoding. Actually, each and every ASP script should start with that one line. Unfortunately, most Classic ASP applications I have seen, did not. With some tragic consequences in most cases.</li>
</ul> 

</p>

<p>The second line in aspLite/aspLite.asp reads: <strong>Option Explicit</strong>. This is questionable. I assume you know that by having this line as the very first line in your ASP script, you are forced to declare variables and you're not allowed to declare them more than once. Even though it helps to keep the risks for overwriting certain values under control, it is not a 100% guarantee. Especially when using VBScript's Execute and ExecuteGlobal statements, Option Explicit does not have any effect. So be careful. Make a very good deal with yourself: always declare variables and keep their naming logical and consistent. That's harder than it sounds. Even though in theory you can declare (dim) the same variables (i, j, rs, counter,... are amongst the more popular variable names) in each and every class, function and sub, they DO overwrite each other. That's no doubt one of the reasons "real" developers never liked VBScript. You never really knew what value (and what type) a VBScript variable held. And if you're not used to that, I can imagine this is driving a developer crazy.</p>

<div class="mt-4 mb-4 p-4  text-bg-dark lead">
I believe that the lack of strictly or strongly typed variables in VBScript has caused it's sudden death back in the early 00s and forced the MicroSoft dev-team to come up with their totally new approach: ASP.NET. It's true that Classic ASP is very prone to a total variable-jungle especially when using lots of include files, not to mention the real disaster scenario when multiple developers had to work on the same code base. It was nearly impossible to work as a team on a Classic ASP application. However, when working alone - like I always did in my career - it is ... feasible as long as you keep some variable-hygiene into account. PHP was - at that time - lousy typed as well. That did not stop PHP from growing and taking over the web development business as from somewhere mid 2000. My 2 cents.
</div>

<p>Right after Option Explicit, 2 ASP files are included.</p>
<p>The first one: <code>&lt;!-- #include file="config.asp"--&gt;</code>
</p>

<p><code>const aspL_path="aspLite"</code> lets you decide where exactly you want the aspLite "engine" in your application.</p>
<p><code>const aspL_debug=true</code> lets you decide whether or not aspLite throws errors. I personally always keep this "true".</p>

<p>The next include file is <code>&lt;!-- #include file="aspForm.asp"--&gt;</code>. That file holds a class, but it is not instantiated. Let's not go into detail right now. Bottom line: it's doing nothing. For now.</p>
	
<p>The next (and last) thing that aspLite/aspLite.asp does is create an instance of aspLite (cls_aspLite):<br>
<code>dim aspL<br>set aspL=new cls_aspLite</code></p>

<div class="alert alert-info">Do you know that it could also have been:<br><code>dim aspL : set aspL=new cls_aspLite</code><br>That would have saved one line of code.</div>

<p>When creating an instance of a VBScript class, the <strong>Private Sub Class_Initialize()</strong> gets executed (if any). Always. And before anything else.</p>
<p>Every line in <strong>Class_Initialize</strong> gets executed every time an instance of <strong>cls_aspLite</strong> instantiates. So we better think twice before we add an infinite loop over there. Ok, bad joke. But each and every line in <strong>Class_Initialize</strong> is thoughtfully written.</p>

<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/classinit.inc"))%></pre>

<p>As we're using <strong>Option Explicit</strong>, we're forced - here also - to declare all variable names. In case of a class - rather than "dim" a variable (even though that also "works" in a way), you can use the private or public keyword. Private variables are available within the class only. Public variables (and their values) can be exposed to the outside world. When working alone on apps, you actually don't really need private values. But as aspLite might at some point in time be used by another developer, I guess I can better do it right, and keep most variables in de aspLite class private.<p>

<p>As the code in <strong>Class_Initialize</strong> is always executed when an instance of cls_aspLite is created, let's have a close look at what happens, line by line.</p>

<p><code>on error resume next</code><br>This basically tells the ASP-compiler to continue processing the lines below, even in case an error is thrown. But you knew that already, didn't you? What you also have to know, is that this statement needs to be repeated in each and every function or sub. In essence, this "resume next" will be reset at the end of <strong>Class_Initialize</strong>. In the next function, snippet or sub, the ASP compiler will - again - stop executing the script in case an error is thrown. Good to know I guess.</p>

<p>
<code>startTime					= Timer()</code><br>
Just because we can and for the fun of it, aspLite holds a timer. startTime will hold the start-time of the script. Let's do this. So far, this page took <strong><%=aspl.printTimer%> milliseconds</strong> to load. That's not much. Having this <code>aspl.printTimer</code> at your fingertips, can help you to isolate badly written code or isolate code that really runs too slow.</p>

<p>
<code>debug						= aspL_debug</code><br>The value for aspL_debug was set in <code>&lt;!-- #include file="config.asp"--&gt;</code>.</p>	

<p>
<code>Response.Buffer				= true</code><br>
Honestly, this is questionable, again. True is the default value anyway. For a reason. If you need to empty the buffer before the ASP script has completely been executed, you can use response.flush and response.clear as often as you wish. As Response.buffer=true is the default value, this line could have been skipped.		
</p>

<p>
<code>Response.CharSet			= "utf-8" 'does not work on IIS5 (Windows 2000 Servers) - comment it out when IIS5 is used!</code><br>	
Questionable. We already made sure utf-8 is our default charset by adding CODEPAGE="65001" in the first line of aspLite/aspLite.asp. As the VBScript comments indicate, this line does not work in IIS5 (Windows 2000 Servers). Hence the "On error resume next" above.
</p>

<p>
<code>Response.ContentType		= "text/html"</code><br>
99% of the output of a web application consists out of text/html. So it's the default content type. It can easily be overwritten if needed though. More about that when discussing the file-serving capabilities of aspLite.
</p>		

<p>These next four lines ensure that browsers do not cache any output by any ASP page in our project. This is crucial. Back in the late 90s, browser caching was considered useful, as internet connections where slow. Today, you really do not want browsers to cache anything, except the things you really want them to cache (cookies, localStorage, etc).<br>
<code>Response.CacheControl		= "no-cache"</code><br>
<code>Response.AddHeader "pragma", "no-cache"</code><br>
<code>Response.Expires			= -1</code><br>
<code>Response.ExpiresAbsolute	= Now()-1</code>
</p>		

<p>		
<code>Server.ScriptTimeout		= 3600</code><br>
I agree, this is quite a tolerant value. Our ASP scripts can run for an hour before timing out. Never ever let an ASP page run for an hour. But in some very rare cases, you may have no other option, like when dealing with large file-transfers or occasional migrations or synchronizations.
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
<code>cacheprefix="aspLite_"</code><br>
</p>

<p>aspLite includes some collections (VBScript dictionaries) and instances of other classes. To be able to check whether they are "nothing", they are set to "nothing" to begin with. It's a trick but very effictive.<br> 		
<code>set plugins			= nothing</code><br>
<code>set p_fso			= nothing</code><br>
<code>set p_JSON			= nothing</code><br>
<code>set p_randomizer	= nothing</code><br>
<code>set p_formmessages	= nothing</code><br>
</p>

<p>
<code>on error goto 0</code><br>
This is technically not needed, as we're at the end of the sub anyway. After this, the On Error Resume Next will not have any effect anymore. However, I prefer to clear possible errors before continuing. That's what this line is for. It "wipes" the VBScript Err-object.
</p>

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
aspl.executeASP(aspl.XMLhttp("https://demo.aspLite.com/ebook/helloworld.txt",false))
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

<p>Finally, browse to</p>

<ul>
<li><a href="http://localhost/?script=1">http://localhost/?script=1</a></li>
<li><a href="http://localhost/?script=2">http://localhost/?script=2</a></li>
<li><a href="http://localhost/?script=3">http://localhost/?script=3</a></li>
</ul>

<p>The one and only default.asp runs 3 different "imported" scripts. Unlike when using include-directives, the imported ASP code is dynamically loaded (or not). You decide which ASP script has to load and which does not. This keeps RAM usage under control. But above all, it facilitates an amazingly well structured ASP code base. For each event you can have your dedicated classes, functions, subs and program flow. And stil you would only be using only one ASP page (default.asp). That one ASP page is able to serve each and every module of your ASP application.</p>

<div class="mt-4 mb-4 p-4  text-bg-success lead">
<p><strong>CDN for Classic ASP/VBScript classes?</strong></p>
<p>The Hello-World script above (https://demo.aspLite.com/ebook/helloworld.txt) could have lead - back in 2000 - to another idea. But it didn't. Anyway. The idea would be to setup a CDN (Content Delivery Network) serving well-written ASP-classes. Pretty much like JavaScript frameworks rely on CDN, Classic ASP/VBScript CDN could have been setup in a very similar way. It would not be the browser loading a CDN file, but the server loading full-blown (compressed) classes to deal with (less) common tasks and functions. THAT's what we needed in 2002. We did not need ASP.NET. Also, bespoke CDN would have been a major step towards Open Source. Still today, it would be a great help for Classic ASP developers. We don't even need a hosting solution. All we need is a couple of megabytes of cloud storage. This idea of CDN for Classic ASP goes beyond the scope of this book. It's something worth experimenting with however.</p>
</div>

<div class="mt-4 mb-4 p-4  text-bg-warning lead">
<p><strong>Be careful with aspl.exec() and aspl.executeASP().</strong></p>
<p>Dynamically importing ASP code with aspl.exec or aspl.executeASP is very powerful. It can be used to create plugins, import remote code, keep your codebase strictly structured, etc. Unfortunately there is no way to trace errors in ASP code that is imported this way. Using regular include files will return appropriate error messages like we're used to and we would always need to during development. The underlying VBScript function <code>executeGlobal</code> is to blame. Make sure to use <code>aspl.exec</code> and <code>aspl.executeASP</code> with care and only for small ASP snippets or well-tested ASP code.</p>
</div>

<a name="getRequest"></a><h4>aspl.getRequest(value)</h4>

<p><code>aspl.getRequest(value)</code> replaces the generic ASP Request object. Unlike the built-in ASP Request Object, it does not throw an error in case a file was uploaded through a form. Pretty much like in the original ASP Request Object, there is a sort order of what exactly is returned: first the form, next the querystring and finally cookies if any.</p>

<h4>aspl.htmlEncode(value)</h4>

<p>Previous example made use of another aspLite function: <code>aspl.htmlEncode</code>. Unlike server.htmlEncode(value), aspl.htmlEncode(value) does not throw an error when value is null.</p>

<div class="alert alert-danger lead">It can't be said enough: ALWAYS htmlencode ANY input from users. I mean: also first names, last names and street numbers. <strong>Anything</strong>. It's the first and most efficient protection against XSS (Cross-Site-Scripting).</div>

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
<p>Example: <code>&lt;input type="date" value=&lt;%=aspl.convertHtmlDate(date)%&gt; /&gt;</code> returns: <input type="date" value=<%=aspl.convertHtmlDate(date)%> /></p>
		
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

<h4>aspl.padLeft(value, totalLength, paddingChar)</h4>
<p>Adds paddingChar left to value until totalLength is reached. Example: <strong><%=aspl.padLeft("5",10,"0")%></strong></p>

<h4>aspl.getFileType(filename)</h4>
<p>Returns the file extension for a given filename. Example: <strong><%=aspl.getFileType("photo.jpeg")%></strong></p>
		
<h4>aspl.log(value)</h4>
<p>aspLite comes with a basic logger: <code>aspl.log("anything")</code> will write "anything" to aspLite/aspLite.log. The exact time of the logging is included as well. This logging feature is very useful as you can always tell what exactly happens to a variable, or when things go wrong. I often use it to debug certain functions.</p>

<h4>aspl.recycleApplication</h4>
<p>Resets an ASP application by opening, saving and closing the global.asa-file.</p>	

<h4>aspl.randomizer</h4>
<p>Randomizer class with 3 functions:
	<ul>
		<li><code>aspl.randomizer.randomText(nmbrchars)</code> returns a random string with a given nmbrchars</li>
		<li><code>aspl.randomizer.randomNumber(start,stop)</code> returns a random number between start and stop</li>
		<li><code>aspl.randomizer.createGUID(length)</code> returns a GUID of a given length</li>				
	</ul></p>
	
<h4>aspl.removeAllCookies</h4>

<p>Removes all cookies</p>

<h4>aspl.printTimer</h4>

<p>Returns the actual execution time of your ASP script.</p>

<h4>aspl.dump(value)</h4>

<p>Sends value to the browser as text/html. aspLite destroys itself and all its plugins right after. In case value includes [aspLite_executionTime], the actual execution time will be added to the output as an HTML comment.</p>	

<h4>Application Caching</h4>

<p>ASP comes with a very interesting object, the Application Object. It was designed to store some application-wide values. In most cases, ASP developers only used it very minimally. They typically stored only very few strings and numbers in the Application Object. However, it was designed to be used a lot heavier than that. It can easily store 10000s of values. When pumped with Arrays, the Application Object is nothing less but a very powerful dictionary or collection object. That's why I built a very tiny layer around it in aspLite.</p>

<p><code>aspl.setcache(name,value)</code> stores an array in the Application Object. Not only the name and the value are stored, also the exact time. Remember the cacheprefix above? It helps not to interfere with other Application variables you may already use in your ASP script.</p>	

<p><code>aspl.getcache(name)</code> returns the Application value for a given name.</p>

<p><code>aspl.getcacheT(name,seconds)</code> returns the Application value for a given name only if it was stored less than x seconds ago.</p>

<p><code>aspl.clearAllCache()</code> clears all Application values for your cacheprefix.</p>

<h3>Textfiles and binaries</h3>

<p>Even though very useful, the above settings and methods are not what you'd call a true framework. They're just a fixed set of configurations and functions you'd use in each and every project. But not more than that.</p>

<p>For years I have assumed that ASP/VBScript was not capable of dealing with (large) binary files (upload, read, write, save, download). At least, that was, without expensive third party COM components. I was wrong all this time. Lewis Moten once created a purely scripted (and free-to-use) ASP/VBScript Upload class. It did and still does a very good job.</p>

<div class="mt-4 mb-4 p-4  text-bg-info lead">
<h4>FileSystemObject vs ADODB.Stream</h4>
<p>While the VBScript FileSystemObject is needed to browse files and folders, aspLite uses <strong>ADODB.Stream</strong> to read, write and save files, both binaries and texts. The FileSystemObject does not support binary files and has major issues with dealing with the UTF-8 (Unicode) charset. ADODB.Stream however, perfectly handles binaries and Unicode characters.</p>
</div>

<h4>aspl.loadText(path)</h4>

<p>Returns the content of a textfile. Path is the relative path to a text-file.</p>

<p><code>aspl.loadText(path)</code> is a great help when separating ASP-code from HTML, thus prevent the so-called spaghetti-ASP-way of coding. You would typically want to isolate all html blocks and snippets in a specific folder and dynamically include them in your ASP script when needed. </p>

<p><strong>Example:</strong></p>

<p>In the root of your application, you have "default.asp".</p>
<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/loadText.inc"))%></pre>

<p>The folder "html" holds default.inc:</p>
<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/defaultreplace.inc"))%></pre>

<p>Congratulations! You just wrote your first code-behind ASP file and your first Classic ASP MasterPage.</p>

<div class="mt-4 mb-4 p-4  text-bg-warning lead">
<h5>Which file-extension should I use for protected and/or code-behind files?</h5>
<p>.inc is a good choice. I also use .resx. Both file types are not served by IIS. In other words, they will never be executed by IIS and they will never be sent nor exposed to the browser. Many other formats will either be sent or executed: .asp, .htm(l), .txt. Be careful with them.</p>
</div>

<h4>aspl.saveText(path,data)</h4>
<p>Saves text (data) to a file (relative path). Unlike the FileSystemObject, this one is 100% UTF-8-proof.</p>

<pre class="alert alert-light">
&lt;!-- #include file="aspLite/aspLite.asp"--&gt;
&lt;%
aspl.saveText "html/textfile.txt", "This is the content of the file (text-only). But utf-8-proof: αναγνώστης"
%&gt;
</pre>

<h4>aspl.saveAsFile(fileName, fileBody)</h4>

<p>Sends text-files (only) to the browser (with a download-prompt). Filename is the default filename, filebody is the content of the file.</p>

<p>At first sight, this is not very useful, as you can send any text to the browser already. However, in some cases, you may want to send text-files (.txt, .html, .css, .js, etc), rather than just text. They will then end up in the visitor's default download-folder.</p>

<pre class="alert alert-light">
&lt;!-- #include file="aspLite/aspLite.asp"--&gt;
&lt;%
aspl.saveAsFile "textfile.txt", "This is the content of the file (text-only)"
%&gt;
</pre>

<h4>aspl.loadBinary(filename)</h4>

<p>Returns the content of a binary file (an image, pdf, zip or any other non-text file). Filename is the relative path to a binary file.</p>

<p>As Classic ASP/VBScript does not have much file-editing options, there is not much use for loading binary files, other than using a lot of RAM in case of large files.</p>

<p>Except perhaps when combined with another aspLite function: </p>

<h4>aspl.saveBinaryData(filename,bytearray)</h4>
<p>Saves a given binary file to filename (absolute path!)</p>
<p>The example below copies html/sample.jpg tot html/copy.jpg.</p>

<pre class="alert alert-light">
&lt;!-- #include file="aspLite/aspLite.asp"--&gt;
&lt;%
dim file : file=aspl.loadBinary("html/sample.jpg")

aspl.saveBinaryData server.mappath("html/copy.jpg"), file
%&gt;
</pre>

<p>Another use I can think of is to copy a file from anywhere to your local drive. The example below downloads (GET) the JPG file on https://demo.aspLite.com/default/html/smallfile.jpg and saves it to your local html folder.</p>

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
<p>Returns custom paths in case a 404-redirect is setup in web.config.</p>

<p>For this to work fine, you need the following web.config file in the root of your application:</p>

<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/webconfig.inc"))%></pre>

<p>This tells IIS to execute /default.asp in case a 404-error (file not found) is thrown by IIS. Luckily, the missing filename is passed on via Request.ServerVariables("query_string"). That's where <code>aspLite.pathinfo</code> grabs the missing file or foldername.</p>

<p>The aspLite demo site comes with some examples:<br><a href="https://demo.aspLite.com/about">https://demo.aspLite.com/about</a> points to a non-existing folder. <code>aspl.pathinfo</code> returns "about" in this case. From there one you can easily serve specific content for the "about" page.</p>

<p>This technique of executing scripts in case of a 404 is available from Windows 2000 (IIS5) onwards. It still works on Windows 2022 servers. But it was never documented or approved by MicroSoft as a valid method. That's a shame, because all this time Classic ASP developers could have made use of this technique to offer user-friendly and SEO-optimized urls. I used this technique in QuickerSite from day one. This has meant a big deal and it has drastically improved SEO for all my hosted Classic ASP customers.</p>

<h3>MSXML2.ServerXMLHTTP and MSXML2.DOMDocument</h3>

<p>ASP ships with some very useful http-callers. In short: from within your ASP/VBScript application, you can call basically any other url on the internet, post to any form out there, next wait for its response, and use that output in your application. The output can be text/html, XML or again, pure binary content. I have often used this technique to read RSS-files (typically done with JavaScript today) and copy entire folders recursively from one server to the other (or to any localhost). These ASP functions are also often used to synchronize data between applications via web services.</p>

<p>Another way to look at these two functions is as an Emergency Exit for ASP/VBScript developers. They can be used to organize excursions out of ASP/VBScript and implement PHP or .NET functionalities that are not (easily) available for Classic ASP developers. The technique I'm referring to is to load specific PHP or ASP.NET pages and have them supercharge your ASP script (create images, zip files, pdf files, etc).</p>

<h4>aspl.XMLhttp(url,binary)</h4>
<p>Returns the output of any url - whether or not binary (true/false).</p>
<p>Let's revisit one of the above examples. The line below downloads a JPG file from demo.aspLite.com and saves it to your local html folder. All in one line of ASP code.</p>

<pre class="alert alert-light">
&lt;!-- #include file="aspLite/aspLite.asp"--&gt;
&lt;%
aspl.SaveBinaryData server.mappath("html/smallfile.jpg"),aspl.XMLhttp("https://demo.aspLite.com/default/html/smallfile.jpg",true)
%&gt;
</pre>

<h4>aspl.XMLdom(url)</h4>
<p>Returns an XML document (url) you can list all elements for. Todays web developers would prefer JSON data to XML. More about JSON and aspLite later.</p>

<div class="mt-4 mb-4 p-4  text-bg-warning lead">

<p><strong>Be careful, XMLhttp and XMLdom perform synchronous requests</strong></p>
<p>Unlike AJAX calls in browsers, these two server-side http-request calls perform synchronous calls. That means that ASP waits for the output of the loaded url before it continues to execute the code below. Therefore, you have to be careful not to load urls that take a long time to execute. That would cause your ASP page to hang.</p>

</div>

<p>An example. The code below shows the 6 latest news items from NYTimes' world-stories-RSS: 
<a href="https://rss.nytimes.com/services/xml/rss/nyt/World.xml" target="_blank">https://rss.nytimes.com/services/xml/rss/nyt/World.xml</a>. The RSS is loaded via <strong>aspL.XMLdom</strong>.</p>

<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/feed.inc"))%></pre>

<p><strong>Live example:</strong></p>

<div class="alert alert-light">
<%aspl("ebook/feed.inc")%>
</div>

<div class="mt-4 mb-4 p-4  text-bg-success lead">
<p><strong>Summary</strong></p>

<p>aspLite comes with quite some powerful functions to deal with text files, binaries, whether they reside on your server or anywhere on the internet. Unlike when using VBScript's FileSystemObject, aspLite ensures that any text converts UTF-8-proof.</p>

<p>Classic ASP's serverhttp-requests can open up a whole new world when integrating with other web-technologies and services.</p>

</div>

<div class="pagebreak"></div>

<h2>How aspLite lead to asplForm</h2>

<p>So far we're still not really using or describing a true framework for Classic ASP/VBScript development. We have only added some useful helper-functions through <code>&lt;!-- #include file="aspLite/aspLite.asp"--&gt;</code>, right?</p>

<p>That is about to change in this chapter. When I was developing aspLite during the first weeks of the Covid19 lockdown, I quickly realised aspLite would soon lead to an AJAX formbuilder.</p>

<h2>asplForm explained</h2>

<p>So far aspLite did not rely on any JavaScript framework. And if you don't plan to use asplForm, you don't need jQuery nor any other JavaScript to use aspLite.</p>

<p>asplForm was the first plugin I developed for aspLite. asplForm needs aspLite.js in the browser. aspLite.js relies on jQuery. That makes aspLite a full stack developer framework. Classic ASP developers always were full stack anyway. They needed HTML, CSS and JavaScript for the front and ASP, VBScript, SQL in the back. On top of that, most Classic ASP developers had to deal with IIS administration, SQL Server management, setting up ftp-servers, mail servers, firewalls, backup solutions, Windows server security, etc. That's what I like a lot about being a Classic ASP developer. It's a colorful job with a lot of diversity and complexity. Classic ASP developers were not only full stack, they were often one man bands, taking care of basically everything that related to building applications for their customers.</p>

<div class="mt-4 mb-4 p-4  text-bg-success lead">

<p><strong>What is asplForm exactly?</strong></p>

<p>asplForm facilitates AJAX web forms. asplForm collects form-fields (VBScript dictionaries), sends them to the browser in JSON-format, and aspLite.js converts it to HTML5 web forms. asplForms bind to any HTML element, regular DIV's in most cases. This technique is the heart of the aspLite framework. Again, asplForm is not only relying on pure Classic ASP/VBScript. There is very important role for JavaScript when converting server-generated JSON to HTML controls.</p>

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

<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/defaultfirst.inc"))%></pre>

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

<p>The <code>form.build</code>-method stopts the execution of the ASP script. That's why it's very important to call it here. If we don't, it would each time add the HTML, CSS and JavaScript to its response. We already have these, so we don't need them again.</p>

<p>Some might think: is it a good idea to load default.asp 4 times (in a row)? Maybe not. Or maybe it doesn't really matter either. The most important reason however is that from now on, it's possible to build a web application based on events and clearly organise your ASP code base, using one ASP page only, performing an unlimited number of different tasks. How you structure it, is up to you. There are several ways to do it. You can even use good old include-files for each asplEvent, or use aspLite's execute methods. </p>

<p><strong>Let's take this one step further</strong></p>

<p>Use this <strong>default.asp</strong> in the root of your ASP application:</p>

<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/default.inc"))%></pre>

<p>Note that we're not loading (including) one particular file. We're loading "whatever" value <code>aspl.getRequest("asplEvent")</code> holds. This way, we dynamically load ASP code into the ASP page's namespace. If a file does not exist, aspLite throws a file-not-found error in case aspL_debug (aspLite/config.asp) is "true".</p>

<p>Next, create <strong>html/default.inc</strong> (create the HTML-folder if needed) with the following HTML, CSS and JavaScript:</p>

<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/defaulthtml.inc"))%></pre>

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

<p>Also, this time the code is organised in a different way. Call it code-behind or anything you like. Default.asp executes ASP code, html/default.inc holds all HTML, CSS and JavaScript and the 3 asp-includes each return an asplForm with some content.</p>

<div class="mt-4 mb-4 p-4  text-bg-info lead">
<p><strong>Why would you reload an aspLite form every x seconds?</strong></p>
<ul>
	<li>To keep an ASP session alive.</li>
	<li>To list alle currently logged on users</li>
	<li>To offer real time game-scores</li>
	<li>...</li>
</ul>
</div>

<p><strong>Can't I just use include-files instead of using aspl.exec()?</strong></p>
<p>Sure. This would have worked as well:</p>

<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/includes.inc"))%></pre>

<p>Be careful though! I moved the line <code>dim form : set form=aspl.form</code> to default.asp and removed it in all 3 *.inc files. If not, we get a "name redefined" error. That error is not thrown when using <code>aspl.exec()</code> - or VBScript's executeGlobal. Not sure why that is the case. Maybe it's a bug in executeGlobal.</p>

<p>As such, "good old" including all different ASP scripts like this will not harm performance nor will it need much more RAM. The only inconvenience is that you have to list all events and "manually" assign the appropriate script to them, like in the example above. There is one other reason why you'd prefer regular includes: better error descriptions! Whenever something goes wrong in ASP code imported by <code>aspl.exec()</code>, all you get is vague error message without a line number! That can be very annoying to debug your ASP code.</p>

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

<p>Whenever you start a new ASP development, always use a modern front-end framework. Do not write your own, use an existing one. Back in the late nineties, ASP developers were used to handcode HTML controls, very often on an HTML table-layout. You could really tell if a website was on Classic ASP/VBScript back in those days. They all looked very similar.</p>

<p>An average customer never cares about which technologies their website uses in back. Customers only care about what they see in the front: an attractive design, good usability, nice colors, does it work on a phone, the icon-set you use, etc. Your application can use the very latest ASP.NET version, if it doesn't have an attractive font-end, your customer won't be happy. On the other hand, if you're using a state-of-the-art front-end framework and well-coded Classic ASP/VBScript in the back, your invoices get paid with a smile.</p>

<p><strong>An example</strong></p>

<p>Let's reuse the default.asp from above in the root:</p>

<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/default.inc"))%></pre>

<p>Add <strong>html/default.inc</strong> and import the Bootstrap CSS and JS libraries via CDN:</p>

<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/bootstrap.inc"))%></pre>

<p>Finally, add <strong>asp/main.inc</strong>:</p>

<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/fields.inc"))%></pre>


<p>As <code>form.field("")</code> returns a VBScript dictionary, the following syntax works as well:</p>

<pre class="alert alert-light">
dim hidden : set hidden=form.field("hidden")
hidden("name")="hidden"
hidden("value")="12345"
</pre>

<p>If all goes well, you end up with a form asking for your name, your birthdate, your email address and some remarks. After you successfully submit, some confirmation message shows up. Congratulations! This is your first (fully responsive) AJAX form using Bootstrap, aspLite and asplForm!
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
	<li>When hitting the submit-button, the form is submitted through and AJAX-call and form.postback will be true</li>
	<li>If form.postback is true, you can process all values from the form.</li>

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

<div class="mt-4 mb-4 p-4  text-bg-primary lead">
<p>This JSON-stream is generated by the JSON class in asplite/asplite.asp. It's a 100% copy of the JSON class I found in ASP-Ajaxed on <a target="_blank" style="color:#FFF" class="link" href="https://github.com/ASP-Ajaxed/asp-ajaxed">https://github.com/ASP-Ajaxed/asp-ajaxed</a>. It was originally developed by Michal Gabrukiewicz. His passing in 2009 was sure one of the reasons for this framework not to become what it could have been. With aspLite, I want to honor Michal for his great work on ASP-Ajaxed. He was one of the best ASP developers ever. I consider aspLite at least 50% Michal's work.</p>

<p>JSON parsers are amongst the most popular Classic ASP repositories on GitHub.</p>

<ul>
	<li><a class="link" style="color:#FFF" target="_blank" href="https://github.com/rcdmk/aspJSON">aspJSON</a> by Ricardo Souza</li>
	<li><a style="color:#FFF" class="link" href="https://github.com/gerritvankuipers/aspjson" target="_blank">aspjson</a> by Gerrit Van Kuipers</li>
</ul>

<p>This illustrates how developers find their way around limitations and shortcomings. MicroSoft never added support for JSON to Classic ASP. But that does not stop ASP developers. We need JSON-functionality to integrate modern JavaScript libraries and frameworks into our applications. So we get it done!</p>
</div>

<h4>form.newline</h4>
<p>Adds a linefeed of 7px (height).</p>

<p>Example with <strong>selectboxes, checkboxes and radio buttons</strong>. Keep both default.asp and html/default.inc. <strong>Change asp/main.inc to:</strong></p>


<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/selectboxes.inc"))%></pre>

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

<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/plain.inc"))%></pre>

<p><strong>Live example:</strong></p>
<div class="asplForm alert alert-light" id="plain"></div>

<p>There are shortcuts for both field-types (plain and script):</p>

<pre class="alert alert-light">
form.write "This adds plain/text or &lt;u&gt;HTML&lt;/u&gt;"
form.writejs "alert('Add JavaScripts');"
</pre>

<div class="mt-4 mb-4 p-4  text-bg-primary lead">

<p><strong>asplForms come with a VIEWSTATE-kind of facility</strong></p>

<p>By default asplForms keep track of form-values and autofills form-controls with user-input. There is no need to initialize them "manually" with request-values.</p>

</div>

<h4>form.alert(alerttype,message)</h4>

<p>This helper-function relies on Bootstrap 5. It returns Bootstrap alerts.</p>

<p>Change asp/main.inc to:</p>

<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/alert.inc"))%></pre>

<p><strong>Live example:</strong></p>
<div class="asplForm alert alert-light" id="alert"></div>

<h4>More complex forms with Bootstrap</h4>
<p>Sometimes you need more flexibility when it comes to placing form-controls.</p>

<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/complex.inc"))%></pre>

<p><strong>Live example:</strong></p>
<div class="asplForm alert alert-light" id="complex"></div>

<p>By setting <code>form.className="row"</code>, we initialize Bootstrap's 12-column grid-system. <code>name.add "containerclass","col-md-6"</code> makes sure the "name" field uses half of the available grid-lines (6). This way you can easily have multiple form-fields on 1 line. The good thing about using Bootstrap for this purpose, is that forms remain 100% responsive.</p>


<h4>form.redirect(asplevent,target,querystring)</h4>
<p>Helper function to navigate between asplForms. asplEvent indicates which ASP code to load, target is the ID of the HTML-control that will be bound to and querystring can be used to add custom parameters.</p>

<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/redirect.inc"))%></pre>

<p><strong>Live example:</strong></p>
<div class="asplForm alert alert-light" id="redirect"></div>

<p>After clicking the Redirect-button, the file redirectNew.inc is loaded:</p>
<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/redirectNew.inc"))%></pre>

<p>After clicking the back-button, the ASP code in redirect.inc is loaded again. If "target" is empty, it will inhert the name of asplEvent, thus load the script "ebook/redirect.inc" in div-id "redirect".</p>

<p>Unlike when using "response.redirect" in Classic ASP, you can "redirect" to multiple divs. In other words: multiple divs can be reloaded.</p>

<p><strong>ebook/redirectmultiple.inc reads:</strong></p>

<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/redirectmultiple.inc"))%></pre>

<p>After hitting the "Redirect"-button, 3 divs are loaded: redirect1, redirect2 and redirect3. They each load the same ASP code (as they have the same asplEvent: redirectnewmultiple).</p>

<p><strong>ebook/redirectnewmultiple.inc reads:</strong></p>
<pre class="alert alert-light"><%=server.htmlEncode(aspl.loadText("ebook/redirectnewmultiple.inc"))%></pre>

<p><strong>Live example:</strong></p>
<div class="asplForm alert alert-light" id="redirectmultiple"></div>

<div id="redirect1"></div>
<div id="redirect2"></div>
<div id="redirect3"></div>

<div class="mt-4 mb-4 p-4 text-bg-success lead">

<p><strong>Summary</strong></p>

<p>asplForms turn aspLite into a modern full-stack developer framework for Classic ASP/VBScript developers. It offers a powerful AJAX web form-builder whether or not making use of the wonderful Bootstrap front-end HTML/CSS front-end framework.</p>

<p>The collaboration between server-side Classic ASP/VBScript and browser-side JavaScript frameworks is very promising and opens up a whole new world to ASP developers. </p>

<p>All-in-all, this is a very first attempt to bring back Classic ASP/VBScript development. I would even call it "embryonal". Especially the JavaScript file asplite/asplite.js needs some attention and further tweaking.</p>

<p>It should not be too complicated to add even more helper-functions and as aspLite will grow over the years, more helper functions will be needed to facilitate even more exciting features. For now however I prefer to keep things compact, simple and very robust. That's what I liked about Classic ASP in the first place.</p>

<p>The next chapters will cover more interesting topics and will include many more ideas to make Classic ASP development fun again, whether or not using aspLite as the preferred ASP/VBScript framework.</p>

</div>

<div class="pagebreak"></div>

<h2>aspLite plugins</h2>
<h3>CDO.message</h3>
<h3>Database</h3>
<h3>Access, SQL Server or MySQL?</h3>
<h3>Encryption</h3>
<h3>Jpeg</h3>
<h3>Uploader</h3>
<h3>Create your own</h3>

<div class="pagebreak"></div>

<h2>Excursions</h2>
<h3>Disconnected recordsets</h3>
<h3>Datatables</h3>
<h3>ChromeASP</h3>
<h3>SheetJS</h3>
<h3>Select2</h3>
<h3>Bootstrap Modal</h3>
<h3>Ajax Autocomplete for jQuery</h3>
<h3>PHP ZipArchive in Classic ASP</h3>

</div>


</body>
</html>