/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */



$(document).ready(() => {
    //$('#run').click(run);
    $("#mobile").html("<b>Test Log</b><br>");
    Office.onReady(function() {
        let mail = Office.context.mailbox.item;
        let subject = mail.subject;
        $("#mobile").html("Office is ready<br>");
        $.get("https://pokeapi.co/api/v2/pokemon/" + subject, function(data, status) {
            if(status=="success") {
                $("#log").html("GET request completed<br>");
                $("#pokemonName").text(data.name.charAt(0).toUpperCase() + data.name.slice(1));
                $("#sprite").attr("src", data.sprites.front_default);

                //get types
                let types = "";
                data.types.sort((a, b) => (a.slot > b.slot) ? 1 : -1);
                data.types.forEach(function(type) {
                    types += "<span class='" + type.type.name + "'>&nbsp;&nbsp;" + type.type.name.toUpperCase() + "&nbsp;&nbsp;</span>";
                });

                $("#type").html(types);

                let h = data.height / 10 //height in meters
                $("#stat").append("<b>Height: </b>" + h + "m <b>Weight: </b>" + data.weight + "kg");

                let abilities = "";
                data.abilities.sort((a, b) => (a.slot > b.slot) ? 1 : -1);
                data.abilities.forEach(function(ability) {
                    abilities += "<li>" + ability.ability.name + "</li>"
                });
                $("#data").html("<b>Abilities: </b><br><ul>")
                $("#data").append(abilities);
                $("#data").append("</li><hr>");

                let moves = "";
                data.moves.sort((a, b) => (a.name > b.name) ? 1 : -1);
                data.moves.forEach(function(move) {
                    moves += "<li>" + move.move.name + "</li>";
                });

                $("#data").append("<b>Moves</b><ul>");
                $("#data").append(moves);
                $("#data").append("</ul><hr>");
            }
        });
    });
});
  
// The initialize function must be run each time a new page is loaded
Office.initialize = (reason) => {
    $('#sideload-msg').hide();
    $('#app-body').show();
};

async function run() {
    /**
         * Insert your Outlook code here
         */
}