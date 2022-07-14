$(document).ready(function () {

    $('a.external').click(function () {
        window.open(this.href);
        return false;
    });

});