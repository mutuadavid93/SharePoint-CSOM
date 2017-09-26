'use strict';

$(function () {
    $('#loadButton').click(loadAndInclude);
    $('#camlQueries').click(camlQueries);
    $('#dataBinding').click(dataBinding);
    $('#createList').click(createList);
    $('#createItem').click(createItem);
    $('#updateListItem').click(updateListItem);
    $('#callToHostWeb').click(callToHostWeb);
    
}); //doument ready

function dataBinding(evt) {
    event.preventDefault(evt);

    // We are Using jsRender Templates

    var curCTX = SP.ClientContext.get_current();
    var list = curCTX.get_web().get_lists().getByTitle("Products");

    // The CamlQuery
    // Query Products Lists for All Items Where Category Equal to 
    // Beverages(ID = 1)
    var query = new SP.CamlQuery();
    query.set_viewXml("<View>" +
                            "<Query>" +
                               "<Where>" +
                                  "<Eq>" +
                                     "<FieldRef Name='Category' LookupId='True' />" +
                                     "<Value Type='Lookup'>1</Value>" +
                                  "</Eq>" +
                               "</Where>" +
                            "</Query>" +
                            "<RowLimit>5</RowLimit>"+
                        "</View>");
    var items = list.getItems(query);
    var itemsArray = curCTX.loadQuery(items, "Include(ID, Title, UnitsInStock, UnitPrice)");
    curCTX.executeQueryAsync(success1, fail1);

    function success1() {
        var message = $('#message');
        message.append("<br />");
        var template = $('#products-template');
        message.append(template.render(itemsArray)); // jsrender used
    } // success1()
    function fail1(sender, args) {
        alert("Call failed in camlQueries(). Error: " + args.get_message());
    }
} // dataBinding()

function camlQueries(event) {
    event.preventDefault();

    var curCTX = SP.ClientContext.get_current();
    var list = curCTX.get_web().get_lists().getByTitle("Products");

    // The CamlQuery
    // Query Products Lists for All Items Where Category Equal to 
    // Beverages(ID = 1)
    var query = new SP.CamlQuery();
    query.set_viewXml("<View>"+
                            "<Query>"+
                               "<Where>"+
                                  "<Eq>"+
                                     "<FieldRef Name='Category' LookupId='True' />"+
                                     "<Value Type='Lookup'>8</Value>"+
                                  "</Eq>"+
                               "</Where>"+
                            "</Query>"+
                        "</View>");
    var items = list.getItems(query);
    curCTX.load(items, "Include(ID, Title, QuantityPerUnit)");
    curCTX.executeQueryAsync(success1, fail1);

    function success1() {
        var message = $('#message');
        var ienum = items.getEnumerator();

        while (ienum.moveNext()) {
            message.append("<br />");
            message.append(ienum.get_current().get_item("QuantityPerUnit"));
        } // while loop
    } // success1()
    function fail1(sender, args) {
        alert("Call failed in camlQueries(). Error: " + args.get_message());
    }

} // camlQueries()

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


// Creating A List in Code
function createList(event) {
    event.preventDefault();

    //var curCTX = SP.ClientContext.get_current();
    //var web = curCTX.get_web();

    var curCTX = SP.ClientContext.get_current();
    var web = curCTX.get_web();

    try {
        // Exception Handling Using ScopeErrorHandling
        var list = null;

        var scope = new SP.ExceptionHandlingScope(curCTX);
        var scopeStart = scope.startScope();

        var scopeTry = scope.startTry();
            list = web.get_lists().getByTitle("Code_Tasks_List");
            curCTX.load(list);
        scopeTry.dispose();

        var scopeCatch = scope.startCatch();
            var lci = new SP.ListCreationInformation();
            // Set the List Properties
            lci.set_title("Code_Tasks_List");
            lci.set_templateType(SP.ListTemplateType.tasks);
            lci.set_quickLaunchOption(SP.QuickLaunchOptions.on);
            // Now Create the List
            var list = web.get_lists().add(lci);
        scopeCatch.dispose();

        var scopeFinally = scope.startFinally();
            // console.log("Finnaly: Try and load the list again if created in catch");
            list = web.get_lists().getByTitle("Code_Tasks_List");
            curCTX.load(list);
        scopeFinally.dispose();

        // End the Parent Scope
        scopeStart.dispose();

       // curCTX.executeQueryAsync(success3, fail3);

        curCTX.executeQueryAsync(function () {
            if (scope.get_hasException() == true) {
                alert("The List Does not exists")
                // So we created it
            }
            else {
                alert("The List Already Exists");
            }
        });

    } catch (Exception) {
        alert(Exception.message);
    }
} // createList()



function createItem(event) {
    event.preventDefault();
    var ctxCur = SP.ClientContext.get_current();
    
    try{
        var list = ctxCur.get_web().get_lists().getByTitle("Code_Tasks_List");

        // Now Create the Item
        var ici = new SP.ListCreationInformation();
        var item = list.addItem(ici);
        item.set_item("Title", "Sample Title");
        item.set_item("AssignedTo", ctxCur.get_web().get_currentUser());
        var due = new Date();
        due.setDate(due.getDate() + 7);
        item.set_item("DueDate", due);
        item.update();

        ctxCur.executeQueryAsync(success, fail);
    } catch (Ex) {
        alert(Ex.message);
    }

    function success() {
        var message = $('#message');
        message.html("Item Created Successfully");
    }

    function fail(sender, args) {
        alert("Call failed in createItem(). Error: " + args.get_message());
    }
}// createItem()


function updateListItem(event) {
    event.preventDefault();
    var ctxCur = SP.ClientContext.get_current();
    var items = null;

    try{
        var list = ctxCur.get_web().get_lists().getByTitle("Code_Tasks_List");

        var query = new SP.CamlQuery();
        query.set_viewXml("<View><RowLimit>1</RowLimit></View>");
        var qitems = list.getItems(query);

        items = ctxCur.loadQuery(qitems);
        ctxCur.executeQueryAsync(success, fail);
    } catch (Ex) {
        alert("Caught Error: "+Ex.message);
    }

    function success() {
        if (items.length > 0) {
            var item = items[0];
            item.set_item("Status", "In Progress");
            item.set_item("PercentComplete", 0.10);
            item.update();
        }

        ctxCur.executeQueryAsync(success1, fail);
    }

    function success1() {
        var message = $('#message');
        message.html("Item Updated Successfully");
    }

    function fail(sender, args) {
        alert("Call failed in updateListItem(). Error: " + args.get_message());
    }
} // updateListItem()


// Access Host Web Resources(i.e. Lists/Libraries) from App Web;
function callToHostWeb(evt) {
    event.preventDefault(evt);

    var hostUrl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));

    var context = SP.ClientContext.get_current();
    var hostContext = new SP.AppContextSite(context, hostUrl);
    var web = hostContext.get_web();

    var list = web.get_lists().getByTitle("Categories");
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml("<View />");
    var qitems = list.getItems(camlQuery);

    var items = context.loadQuery(qitems);
    context.executeQueryAsync(success, fail);

    function success() {
        var message = $('#message');
        message.text("Categories in the Host Web List: ");
        message.append("<br />");

        // Loop through ListItems
        $.map(items, function (value, index) {
            message.append(value.get_item("Title"));
            message.append("<br />");
        }); // Each Loop
    }

    function fail(sender, args) {
        alert("Call failed in callToHostWeb(). Error: " + args.get_message());
    }

    // Script FUnction to get Query String Parameter e.g. SPHostUrl
    function getQueryStringParameter(paramToRetrieve) {
        var params = document.URL.split("?")[1].split("&");
        var strParams = "";
        for (var i = 0; i < params.length; i = i + 1) {
            var singleParam = params[i].split("=");
            if (singleParam[0] == paramToRetrieve)
                return singleParam[1];
        } // for loop
    } // getQueryStringParameter()
} // callToHostWeb()