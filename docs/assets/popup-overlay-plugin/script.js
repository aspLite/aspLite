
(function(a){a(document).ready(function(){a("html").hasClass("is-builder")||(a(".mbr-popup").on("show.bs.modal",function(c){var b=a(this);setTimeout(function(){a(".modal-backdrop").css({"background-color":b.attr("data-overlay-color"),opacity:b.attr("data-overlay-opacity")})},0)}),a(".mbr-popup").on("hide.bs.modal",function(c){a(".modal-backdrop").css("opacity",0)}))})})(jQuery);
