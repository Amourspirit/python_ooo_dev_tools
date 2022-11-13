$(document).ready(function () {

    // Process true external links.
    // Links that are on a different domain from the docs.
    // Any links that are have the same domain as the docs
    // will not open in a new window.
    $('a.external').filter(function () {
        return this.href.indexOf(location.origin) != 0;
    }).click(function () {
        window.open(this.href);
        return false;
    });

    // Process links that are marked as external
    // from extensions such as sphinx.ext.extlinks
    // and change their class.
    $('a.external').filter(function () {
        return this.href.indexOf(location.origin) === 0;
    }).removeClass('external').addClass('internal')

    // add link to open figure images in a new window for full view.
    $('figure.diagram img, figure.screen-shot img').css('cursor', 'pointer').click(function () {
        window.open(this.src);
        return false;
    });
});