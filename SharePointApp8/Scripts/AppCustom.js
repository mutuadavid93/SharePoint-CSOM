'use strict';

$(function () {
    // Creating Document Libraries
    $('#createLibrary').click(createLibrary);
    $('#uploadButton').click(uplaodDocumentFile);
    
    
    $('#DeleteList').click(function (evt) {
        evt.preventDefault();

        var listToDelete = $('#listToDelete').val();
        DeleteList(listToDelete);
    });
});

function createLibrary(event) {
    event.preventDefault();

    var context = SP.ClientContext.get_current();
    var web = context.get_web();

    var lci = new SP.ListCreationInformation();
    lci.set_title("Code_Document_Library");
    lci.set_templateType(SP.ListTemplateType.documentLibrary);
    var list = web.get_lists().add(lci);

    // Now Add Fields
    list.get_fields().addFieldAsXml("<Field Type=\"Number\" DisplayName=\"Year\" Min=\"2000\" Max=\"2100\" Decimal=\"0\" Name=\"Year\"  />",
        true, SP.AddFieldOptions.defaultValue);
    list.get_fields().addFieldAsXml("<Field Type=\"User\" DisplayName=\"Coordinator\" Name=\"Coordinator\" List=\"UserInfo\" ShowField=\"ImnName\" UserSelectionMode=\"PeopleOnly\" UserSelectionScope=\"0\"  />",
        true, SP.AddFieldOptions.defaultValue);

    context.executeQueryAsync(success, fail);

    function success() {
        var message = $('#message').text("Document Library Created");
    }

    function fail(sender, args) {
        alert("Call failed in createLibrary(). Error: " + args.get_message());
    }
}; // createLibrary()




// Delete A list
function DeleteList(listValue) {
    var context = new SP.ClientContext();
    var web = context.get_web();
    var list = web.get_lists().getByTitle(listValue);
    list.deleteObject(); // Delete the created list from the site  
    context.executeQueryAsync(ondeletesuccess, ondeletefailed);
}

function ondeletesuccess() {
    alert("List deleted successfully"); // on success bind the results in HTML code  
}

function ondeletefailed(sender, args) {
    alert('Delete Failed' + args.get_message() + '\n' + args.get_stackTrace()); // display the errot details if deletion failed  
}


// Upload Document File into a Document Library
function uplaodDocumentFile(event) {
    event.preventDefault();

    // Does Browser SUpport HTM5 File API?
    if (!window.FileReader) {
        alert("The Browser you're riding on Hates File Upload API");
        return;
    }

    var context = SP.ClientContext.get_current();
    var list = context.get_web().get_lists().getByTitle("Code_Document_Library");

    // get Document( NB: Here Use JavaScript to get The File Not jQuery )
    var elem = document.getElementById('uploadInput');
    var file = elem.files[0];
    var parts = elem.value.split("\\");
    var fileName = parts[parts.length - 1];

    // Now Reader
    var reader = new FileReader();
    reader.onload = function (e) {
        success1(e.target.result, fileName);
    }
    reader.onerror = function (e) {
        alert(e.target.error);
    }
    reader.readAsArrayBuffer(file);

    function success1(buffer, fileName) {
        // FileBuffer needs to be 64bit Endoded Byte Array.
        var bytes = new Uint8Array(buffer);
        var content = new SP.Base64EncodedByteArray();
        for (var b = 0; b < bytes.length; b++) {
            content.append(bytes[b]);
        }

        // upload Document
        var fci = new SP.FileCreationInformation();
        fci.set_content(content);
        fci.set_overwrite(true);
        fci.set_url(fileName);
        var file = list.get_rootFolder().get_files().add(fci);

        // Create Other Metadata at the Same time
        var item = file.get_listItemAllFields();
        item.set_item("Year", new Date().getFullYear());
        item.set_item("Coordinator", context.get_web().get_currentUser());
        item.update();

        context.executeQueryAsync(success2, fail);
    }


    function success2() {
        var message = $('#message').text("Document uploaded");
    }

    function fail(sender, args) {
        alert("Call failed in usingLoad(). Error: " + args.get_message());
    }
} // uplaodDocumentFile();