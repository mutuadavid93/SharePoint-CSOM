'use strict';

$(function () {
    $('#loadButton').click(loadAndInclude);
    $('#camlQueries').click(camlQueries);
    $('#dataBinding').click(dataBinding);
    $('#createList').click(createList);
    $('#createItem').click(createItem);
    $('#updateListItem').click(updateListItem);
    $('#managedMetaDataCreation').click(managedMetaDataCreation);
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
// NB: You Must Give the App Permission from "Permissions Tab" in 
// AppManifest.xml
// NB: Never grant Full-Control Permission in your APP.
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



/// MANAGED METADATA Taxonomy
// Refers to a hierarchical collection of terms that can be defined, then used in 
// an item in SharePoint. 

function managedMetaDataCreation(event) {
    event.preventDefault();
    var lcid = _spPageContextInfo.currentLanguage;

    var context = SP.ClientContext.get_current();
    var web = context.get_web();
    var session = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
    var store = session.get_termStores().getByName("Managed Metadata Service");
    var group = store.createGroup("JavaScript", "03F5F03F-CD25-466E-8C2A-E2C22C42ED75");
    var set = group.createTermSet("Projects", "6E5B7101-24A2-49A6-A031-648F389D8819", lcid);
    set.createTerm("Penske", lcid, "E1B08A5A-D815-4F74-91A5-53D730EE93A1");
    set.createTerm("Manhattan", lcid, "A32D4ECD-82C1-46AA-B32D-BCBF688AE29A");
    set.createTerm("Alan Parsons", lcid, "89B38D86-09CC-4A60-980F-24F35FE900ED");
    context.executeQueryAsync(success, fail);

    function success() {
        var message = jQuery("#message");
        message.text("Terms added");
    }

    function fail(sender, args) {
        alert("Call failed. Error: " +
            args.get_message());
    }
}// managedMetaDataCreation

