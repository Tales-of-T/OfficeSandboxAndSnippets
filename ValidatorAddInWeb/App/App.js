/* Common app functionality */
"use strict";
var App;
(function (App) {
    function initialize() {
        $('body').append('<div id="notification-message">' +
            '<div class="padding">' +
            '<div id="notification-message-close"></div>' +
            '<div id="notification-message-header"></div>' +
            '<div id="notification-message-body"></div>' +
            '</div>' +
            '</div>');
        $('#notification-message-close').click(function () {
            $('#notification-message').hide();
        });
    }
    App.initialize = initialize;
    function showNotification(header, text) {
        $('#notification-message-header').text(header);
        $('#notification-message-body').text(text);
        $('#notification-message').slideDown('fast');
    }
    App.showNotification = showNotification;
})(App || (App = {}));
