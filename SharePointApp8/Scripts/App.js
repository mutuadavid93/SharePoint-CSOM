'use strict';

$(function () {
    $('#loadButton').click(loadAndInclude);
}); //doument ready

function loadAndInclude(event) {
    event.preventDefault(); // Alter Default Refresh Behaviour

    var curCTX = SP.ClientContext.get_current();
    var lists = curCTX.get_web().get_lists();

    // Retrieve List Titles together with Each List's Title Field.
    // Nested Includes
    curCTX.load(lists, "Include(Title, Fields.Include(Title))");

    curCTX.executeQueryAsync(success, fail);

    function success() {
        var message = $('#message');
        var lenum = lists.getEnumerator();
        while (lenum.moveNext()) {
            var list = lenum.get_current();
            message.append("<br />");
            message.append(list.get_title());

            // Go One Step Deep Inside Each List's Fields Collection
            var fenum = list.get_fields().getEnumerator();
            var x = 0;
            while (fenum.moveNext()) {
                var field = fenum.get_current();
                message.append("<br />&nbsp;&nbsp;&nbsp;&nbsp;");
                message.append(field.get_title());

                if(x++ > 3) break; // return only 5 items
            }
        }
    }

    function fail(sender, args) {
        alert("Call failed in loadAndInclude(). Error: " + args.get_message());
    }
} // loadAndInclude()