# ASP-VBScript-Framework

This framework can help to develop better ASP/VBScript applications. 

Install as a website or virtual directory in IIS.

Preview: https://aspfw.quickersite.com

## Highlights:

* split html/css/javascript and asp
* conditionally load asp files (rather than include them 'all')
* charset utf-8
* .net-alike codebehind approach
* light-weight and fast ASP classes
* write your own plugins

## Less is more

The framework itself is only 2 files (asp.asp and its corresponding config.asp file). 
The only thing this framework really does is wrap useful ASP components in reusable functions and classes:

* file streaming (text + binary)
* ServerXMLHTTP & DOMDocument
* application caching
* generic ASP/VBScript functions

## Plugins

The following plugins are included already. But they are not necessary to use the framework i.e. they can be ignored 
and when not used, they will not affect the performance of your application in ANY way.

* connector for Access databases (OLEDB)
* mailer for CDO.Message
* jpg resize/recolor (.net required)
* json read/write
* text/number randomizing
* rss feed reader

This ASP/VBScript framework does not include or favorize frontend CSS/JavaScript libraries. But just for the fun of it, 
I added a Bootstrap sample and some AJAX examples each using jQuery. 

## Roadmap

* make file (and complete folder?) uploads in ASP classic easier (core)
* keep security high (protect against sqli, csrf, brute force attacks, cookie theft, etc) (core)
* facilitate multilingual websites and applications (core)
* implement one or more WYSIWYG editors (as plugins)
