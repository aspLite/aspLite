
(function(b){b(document).ready(function(){b("html").hasClass("is-builder")||(b(".mbr-popup").on("shown.bs.modal",function(a){a=b(this).find(".mbr-embedded-video");-1!==a.attr("src").indexOf("autoplay=1")?a.attr("data-src",a.attr("src")):a.attr("data-src")&&a.attr("src",a.attr("data-src"));a.height(a.width()*parseInt(a.attr("height")||315)/parseInt(a.attr("width")||560))}),b(".mbr-popup").on("hidden.bs.modal",function(a){a=b(this).find(".mbr-embedded-video");a.attr("src",a.attr("src").replace("autoplay=1",
""))}))})})(jQuery);
