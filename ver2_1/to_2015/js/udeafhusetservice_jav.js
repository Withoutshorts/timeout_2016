$(document).ready(function () {

    setTimeout(function () {
        CloseWindow()
    }, 7000);

});

function CloseWindow() {
    window.top.close();
}


