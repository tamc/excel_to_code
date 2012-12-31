$(document).ready(function() {

    $('table.cells td').on('click', function(event) {
      window.location.hash = this.id.substring(1);
    });

    var highlight = function(reference) {
      $('table.cells td').removeClass('selected');
      c = $("#c"+reference)
      c.addClass('selected');
    };

    var showFormula = function(reference) {
      c = $("#c"+reference)
      $('#selectedcell').html(reference);
      $('#selectedformula').html(c.data('formula'));
      $('#selectedvalue').html(c.data('value'));
    }

    $(window).on('hashchange', function(event) {
      reference = window.location.hash.substring(1);
      highlight(reference);
      showFormula(reference);
    });

    if(window.location.hash == "") {
      window.location.hash = $('table.cells td').first()[0].id.substring(1);
    } else {
      $(window).trigger('hashchange'); // Doesn't happen on page load
    };


});
