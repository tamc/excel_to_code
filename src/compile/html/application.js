$(document).ready(function() {

    var setSizes = function() {
      $('#worksheet').height($(window).height() - $('#top').height() - $('#jumpbar').height());
    }

    $(window).resize(setSizes);
    setSizes();

    var scrollTo = function(element) {
      var element = $(element);
      var worksheet = $('#worksheet');
      worksheet.animate({ 
          scrollTop: element.offset().top + (element.height()/2) - (worksheet.height()/2), 
          scrollLeft: element.offset().left + (element.width()/2) - (worksheet.width()/2)
      }, 500);
    };

    $('table.cells td').on('click', function(event) {
      window.location.hash = this.id.substring(1);
    });

    var highlight = function(reference) {
      $('table.cells td').removeClass('selected');
      c = $(".c"+reference)
      c.addClass('selected');
    };

    var showFormula = function(reference) {
      c = $(".c"+reference)
      $('#selectedcell').html(reference);
      $('#selectedformula').html(c.data('formula'));
      $('#selectedformula a').on('click', function(event) { scrollTo(this) });
    }

    $(window).on('hashchange', function(event) {
      reference = window.location.hash.substring(1);
      highlight(reference);
      showFormula(reference);
    });

    if(window.location.hash == "") {
      window.location.hash = "A1"
    } else {
      $(window).trigger('hashchange'); // Doesn't happen on page load
    };

    scrollTo($('.selected'));
});
