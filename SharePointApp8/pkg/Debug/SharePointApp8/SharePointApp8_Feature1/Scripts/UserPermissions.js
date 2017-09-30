'use strict';

$(function () {
    $('#userPermiSions').click(userPermiSions);
    $('#getProfileProperties').click(getProfileProperties);
    $('#searchFunc').click(searchFunc);
    
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
// NB: User Profiles are Stored outside the App. Thus configure AppManifest for Read Perm
function getProfileProperties(evnt) {
    evnt.preventDefault();

    var context = SP.ClientContext.get_current();
    var manager = new SP.UserProfiles.PeopleManager(context);
    var profile = manager.getMyProperties();

    context.load(profile, "DisplayName", "UserProfileProperties");
    context.executeQueryAsync(success, fail);

    function success() {
        var message = $('#message');
        message.text("User Profile Properties for " + profile.get_displayName());
        message.append("<br />");

        // iterate thro profile returning only non-null properties
        var props = profile.get_userProfileProperties();
        for (var key in props) {
            var value = props[key];
            if (value.length > 0) {
                message.append(key + " : " + value);
                message.append("<br />");
            }
        }
    }

    function fail(sender, args) {
        alert("Call failed in getProfileProperties(). Error: " + args.get_message());
    }

} // getProfileProperties()


// Search
// Perform a Keyword Query Search
function searchFunc(event) {
    event.preventDefault();

    var context = SP.ClientContext.get_current();

    try {
        var queryText = "boxes";
        var query = new Microsoft.SharePoint.Client.Search.Query.KeywordQuery(context);
        query.set_queryText(queryText);

        var exec = new Microsoft.SharePoint.Client.Search.Query.SearchExecutor(context);
        var results = exec.executeQuery(query);

        context.executeQueryAsync(success, fail);
    } catch (Exception) {
        alert(Exception.message);
    }

    function success() {
        var message = $('#message');
        message.text("Search Results for \"" + queryText + "\"");
        message.append("<br />");
        
        var rows = results.m_value.ResultTables[0].ResultRows;
        $.each(rows, function (index, value) {
            message.append(value.Title + ": " + value.Path);
            message.append("<br />");
        });
    }

    function fail(sender, args) {
        alert("Call failed in searchFunc(). Error: " + args.get_message());
    }
} // searchFunc()
