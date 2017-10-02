'use strict';

$(function () {
    $('#userPermiSions').click(userPermiSions);
    $('#getProfileProperties').click(getProfileProperties);
    $('#searchFunc').click(searchFunc);
    $('#webServiceAccess').click(webServiceAccess);
    
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
        
        // Return the Item Title and Associated Link to View It's Details.
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


// Access a Web Service
// You need to add the Remote URL to the AppManifest.xml
function webServiceAccess(evnt) {
    evnt.preventDefault();

    var context = SP.ClientContext.get_current();

    var request = new SP.WebRequestInfo();
    request.set_url("http://services.odata.org/northwind/northwind.svc/Categories?$format=json");
    request.set_method("GET");
    var response = SP.WebProxy.invoke(context, request);

    context.executeQueryAsync(success, fail);

    function success() {
        // NB: WebProxy Only Calls Success not Fail
        if (response.get_statusCode() == 200) {
            var categories = JSON.parse(response.get_body());

            var message = $('#message');
            message.text("Categories in the remote Northwind Web Service: ");
            message.append("<br />");

            // iterate thro the categories
            $.each(categories.value, function (index, item) {
                message.append(item.CategoryName);
                message.append("<br />");
            }); // each Loop
        } else {
            var errorMessage = response.get_body();
            alert(errorMessage);
        }
    }

    function fail(sender, args) {
        alert("Call failed in webServiceAccess(). Error: " + args.get_message());
    }
} // webServiceAccess()