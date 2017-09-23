'use strict';

$(function () {
    $('#loadButton').click(loadAndInclude);
}); //doument ready

function loadAndInclude(event) {
    event.preventDefault(); // Alter Default Refresh Behaviour

    var curCTX = SP.ClientContext.get_current();
    var lists = curCTX.get_web().get_lists();

    curCTX.load(lists, "Include(Title)");
    curCTX.executeQueryAsync(success, fail);

    function success() {
        var message = $('#message');
        var lenum = lists.getEnumerator();
        while (lenum.moveNext()) {
            message.append("<br />");
            message.append(lenum.get_current().get_title());
        }
    }

    function fail(sender, args) {
        alert("Call failed in loadAndInclude(). Error: " + args.get_message());
    }
} // loadAndInclude()