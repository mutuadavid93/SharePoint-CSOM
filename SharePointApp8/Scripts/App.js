'use strict';

$(function () {
    $('#loadButton').click(loadAndInclude);
    $('#camlQueries').click(camlQueries);
    $('#dataBinding').click(dataBinding);
    $('#createList').click(createList);
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

    var curCTX = SP.ClientContext.get_current();
    var web = curCTX.get_web();

    try {
        var lci = new SP.ListCreationInformation();

        // Set the List Properties
        lci.set_title("Code_Tasks_List");
        lci.set_templateType(SP.ListTemplateType.tasks);
        lci.set_quickLaunchOption(SP.QuickLaunchOptions.on);

        // Now Create the List
        var list = web.get_lists().add(lci);

        curCTX.executeQueryAsync(success3, fail3);
    } catch (Exception) {
        alert(Exception.message);
    }

    function success3() {
        var message = $('#message');
        message.text("List Created Successfully.");
    }

    function fail3(sender, args) {
        alert("Call failed in createList(). Error: " + args.get_message());
    }
} // createList()