$(document).ready(function() {
    $('#worksheet').on('change', function(event) {
      window.location = event.target.value;
    });

    $('table.cells td').on('click', function(event) {
      window.location.hash = this.id;
      $('#selectedcell').html(this.id);
      $('#selectedformula').html($(this).data('formula'));
      $('#selectedvalue').html($(this).data('value'));
      $('table.cells td').removeClass('selected');
      $(this).addClass('selected');
    });    
    $(window).on('hashchange', function(event) {
      $(window.location.hash).trigger('click');
    });
    if(window.location.hash == "") {
      $('table.cells td').first().trigger('click');
    } else {
      $(window.location.hash).trigger('click');
    }
});
