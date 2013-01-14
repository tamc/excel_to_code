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

  
    var singleReference = /^[A-Z]+[0-9]+$/
    var rangeReference = /^([A-Z]+[0-9]+):([A-Z]+[0-9]+)$/
    var cellsFromReference = function(reference) {
      if(singleReference.test(reference)) {
        return {start: reference, end: reference};
      } else if(m = rangeReference.exec(reference)) {
        return {start: m[1], end: m[2]};
      } else {
        return {start: 'A1', end: 'A1'};
      }
    };

    $('#worksheet').append("<div id='highlight'>&nbsp;</div>");
    var highlight_div = $('#highlight');

    var highlight = function(start_reference, end_reference) {
      if(end_reference == undefined) {
        end_reference = start_reference;
      }
      s = $(".c"+start_reference);
      e = $(".c"+end_reference);
      w = $('#worksheet');
      so = s.position(); // Returns relative to #worksheet, but varies based on overflow
      eo = e.position();
      sl = so.left + w.scrollLeft(); // We are setting relative to #worksheet, but should not vary with overflow
      st = so.top + w.scrollTop();
      w = eo.left - so.left + e.width();
      h = eo.top - so.top + e.height();
      highlight_div.animate({left: sl, top: st, width: w, height: h}, 1000);
    };

    var showFormula = function(reference) {
      c = $(".c"+reference)
      $('#selectedcell').html(reference);
      $('#selectedformula').html(c.data('formula'));
      $('#selectedformula a').on('click', function(event) { scrollTo(this) });
    }

    $(window).on('hashchange', function(event) {
      cells = cellsFromReference(window.location.hash.substring(1));
      highlight(cells.start, cells.end);
      showFormula(cells.start);
    });

    if(window.location.hash == "") {
      window.location.hash = "A1"
    } else {
      $(window).trigger('hashchange'); // Doesn't happen on page load
    };

    //scrollTo($('.selected'));
});
