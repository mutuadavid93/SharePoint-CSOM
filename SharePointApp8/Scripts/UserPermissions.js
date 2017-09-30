'use strict';

$(function () {
    $('#userPermiSions').click(userPermiSions);
});

// User Permissions:
// Check the Current USer Permissions at Both the Site and a Document Lib Level
function userPermiSions(event) {
    event.preventDefault();

    var context = SP.ClientContext.get_current();
    var web = context.get_web();
    var list = web.get_lists().getByTitle("Code_Document_Library");

    context.load(list, "EffectiveBasePermissions");

    // Check for the Web
    // i.e. Use a Mask
    var mask = new SP.BasePermissions();
    mask.set(SP.PermissionKind.manageLists);
    var manageLists = web.doesUserHavePermissions(mask);

    context.executeQueryAsync(success, fail);

    function success() {
        // check if user has AddItems Permission
        var perm = list.get_effectiveBasePermissions();
        var addListItems = perm.has(SP.PermissionKind.addListItems);

        var message = $('#message');
        message.text("Manage Lists: " + manageLists.get_value());
        message.append("<br />");
        message.append("Add List Items: "+addListItems);
    }

    function fail(sender, args) {
        alert("Call failed in userPermiSions(). Error: " + args.get_message());
    }

} // userPermiSions()


// User Profiles

