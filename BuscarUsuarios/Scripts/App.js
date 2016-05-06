'use strict';

var clientContext = SP.ClientContext.get_current();

//Funcion encargada de realizar las busquedas
function searchUsers() {
    // limpiar las cajas
    $("#users").html("");
    $("#profile").html("");

    // obtienes los criterios de busqueda
    var userName = $('#accountName').val();
    var name = $('#name').val();
    var department = $('#department').val();
    var searchCriteria = [];

    //Guardas los criterios de busqueda si existen, en un array para realizar la busqueda
    if (userName.length > 0) {
        searchCriteria.push("AccountName:" + userName);
    }
    if (name.length > 0) {
        searchCriteria.push("(FirstName:" + name + " OR LastName:" + name + " OR PreferredName:" + name + ")");
    }
    if (department.length > 0) {
        searchCriteria.push("Department:" + department);
    }

    var queryText = "";
    $.each(searchCriteria, function (index) {
        if (queryText.length > 0) {
            queryText = queryText + " AND ";
        }
        queryText = queryText + this;
    });

    if (queryText.length == 0) {
        return;
    }

    var query = new Microsoft.SharePoint.Client.Search.Query.KeywordQuery(clientContext);
    query.set_queryText(queryText);
    query.set_sourceId("B09A7990-05EA-4AF9-81EF-EDFAB16C4E31");

    var searchExecutor = new Microsoft.SharePoint.Client.Search.Query.SearchExecutor(clientContext);
    var results = searchExecutor.executeQuery(query);

    //Primera funcion si la busqueda tiene exito, segunda función cuando la búsqueda no tiene exito
    clientContext.executeQueryAsync(function () {
        $("#search").hide();
        $("#results").show();
        $('#accountName').val("");
        $('#name').val("");
        $('#department').val("");

        if (results.m_value.ResultTables[0].ResultRows.count < 1) {
            $("#users").append("No se encuentran usuario para los criterios de búsqueda.");
            return;
        }

        $("#users").append("<h1>Users</h1>");
        $.each(results.m_value.ResultTables[0].ResultRows, function () {
            $("#users").append("<div>" +
            "<a href='#' onclick='displayProfile(this)' data-username='" + this.AccountName + "'>" + this.PreferredName + "</a></div>");
        });

    }, function () {

        alert("Error al realizar la búsqueda.");

    });

}

function displayProfile(link) {
    var requiredProperties = ["AccountName", "FirstName", "LastName", "PreferredName", "Department", "Company"];
    var username = $(link).attr("data-username");
    var peopleManager = new SP.UserProfiles.PeopleManager(clientContext);
    var profilePropertiesRequest = new SP.UserProfiles.UserProfilePropertiesForUser(clientContext, username, requiredProperties);
    var profileProperties = peopleManager.getUserProfilePropertiesFor(profilePropertiesRequest);

    clientContext.load(profilePropertiesRequest);

    clientContext.executeQueryAsync(function () {
        $("#profile").html("<div>" +
        "<h1>" + profileProperties[3] + "</h1>" +
        "<p>Nombre: " + profileProperties[1] + "</p>" +
        "<p>Apellidos: " + profileProperties[2] + "</p>" +
        "<p>Departamento: " + profileProperties[4] + "</p>" +
        "<p>Compañia: " + profileProperties[5] + "</p>");

    }, function () {
        alert("Error retrieving user profile.");
    });

}

function showSearch() {
    $("#search").show();
    $("#results").hide();
}

$(document).ready(function () {
    $('#submitSearch').click(searchUsers);
    $("#results").hide();
});