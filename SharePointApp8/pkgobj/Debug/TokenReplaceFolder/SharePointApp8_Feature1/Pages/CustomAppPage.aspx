﻿<%@ Page language="VB" MasterPageFile="~masterurl/default.master" Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<asp:Content ContentPlaceHolderId="PlaceHolderAdditionalPageHead" runat="server">
    <script type="text/javascript" src="../Scripts/jquery-1.9.1.min.js"></script>
    <script type="text/javascript" src="/_layouts/15/sp.runtime.js"></script>
    <script type="text/javascript" src="/_layouts/15/sp.js"></script>
    <script type="text/javascript" src="/_layouts/15/SP.Taxonomy.js"></script>
    <script type="text/javascript" src="/_layouts/15/SP.UserProfiles.js"></script>
    <script type="text/javascript" src="/_layouts/15/SP.Search.js"></script>

    <%--Custom JS Scripts--%>
    <script type="text/javascript" src="../Scripts/AppCustom.js"></script>
    <script type="text/javascript" src="../Scripts/UserPermissions.js"></script>

    <meta name="WebPartPageExpansion" content="full" />


</asp:Content>

<asp:Content ContentPlaceHolderId="PlaceHolderMain" runat="server">
    <div>
        <p id="message"></p>
        <br />
        
        <p><button role="button" id="createLibrary">createDocLibrary</button></p>
        <p><button role="button" id="userPermiSions">User Permissions</button></p>
        <p><button role="button" id="getProfileProperties">User Profile</button></p>
        <p><button role="button" id="searchFunc">Search</button></p>
        <p><button role="button" id="webServiceAccess">Access WebService</button></p>
        
        
        <p>
            <input type="text" id="listToDelete" /><button role="button" id="DeleteList">Delete List</button>
        </p>

        <p>
            <input id="uploadInput" type="file" />
            <input id="uploadButton" type="button" value="Upload Document" />
         </p>
    </div>

    <WebPartPages:WebPartZone runat="server" FrameType="TitleBarOnly" ID="full" Title="loc:full" />
</asp:Content>
