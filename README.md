# aspLite

aspLite is a framework for ASP/VBScript developers. It can help to develop better ASP/VBScript applications. 

More info: https://asplite.com

## Installation:

Install as a website or virtual directory in IIS. The actual framework is located in /asplite/asplite.asp. The path /asplite can be changed in /asplite/config.asp. If you prefer to use aspLite in a subfolder /mypath/asplite, make sure to update config.asp correspondingly.

## Highlights:

* split html/css/javascript and asp
* conditionally load asp files (rather than include them 'all')
* charset utf-8
* .net-alike codebehind approach
* light-weight and fast ASP classes
* write your own plugins

## Less is more

aspLite consists out of 2 files only (asplite.asp & config.asp file). The only thing this framework really does is wrap useful ASP components in reusable functions and classes:

* file streaming (text + binary)
* ServerXMLHTTP & DOMDocument
* application caching
* some basic AJAX functionality
* some generic ASP/VBScript functions

## Plugins

The following plugins are included already - but they are not necessary to use the framework i.e. they can be ignored. 
When no plugins are used, they will not affect the performance of your application in ANY way.

* connector for Access databases (OLEDB)
* mailer for CDO.Message
* jpg resize/recolor (.net required)
* json read/write
* text/number randomizing
* rss feed reader

aspLite does not include or favorize frontend CSS/JavaScript libraries. But just for the fun of it, 
I added a Bootstrap sample and some AJAX examples each using jQuery. 

## Roadmap for aspLite

* make file (and complete folder?) uploads in ASP classic easier (core)
* keep security high (protect against sqli, csrf, brute force attacks, cookie theft, etc) (core)
* facilitate multilingual websites and applications (core)
* implement one or more WYSIWYG editors (as plugins)
