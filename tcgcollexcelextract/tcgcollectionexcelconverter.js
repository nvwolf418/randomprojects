/*
*CHROME OR FIREFOX ONLY
* Steps:
* 1) Go to the tcgplayer collection link and sort it
* 2) Copy everything between dotted lines below that you want to get below, 3 options(foils, no foils, everything)
* 3) Select address bar where tcgplayer url is and press ctrl + a, then type- 
*     javascript:
* 4) press ctrl + v
* 5) press enter, download should start! 
* NOTE
*/



---------------------------@@@@EXTRACT THE ENTIRE COLLECTION@@@@----------------------------------------------


function fnExcelReport() {
    var tab_text = "<table border='2px'><tr bgcolor='#87AFC6'>";
    var textRange;
    var j = 0;
    tabs = document.getElementsByClassName('CollectionTable tablesorter'); 
    tab = tabs[0];
    for (j = 0; j < tab.rows.length; j++) {
        rowInnerHtml = tab.rows[j].innerHTML;

        tab_text = tab_text + rowInnerHtml + "</tr>";
    }

    tab_text = tab_text + "</table>";
    tab_text = tab_text.replace(/<A[^>]*>|<\/A>/g, "");
    tab_text = tab_text.replace(/<img[^>]*>/gi, "");
    tab_text = tab_text.replace(/<input[^>]*>|<\/input>/gi, "");

	sa = window.open('data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,' + encodeURIComponent(tab_text));

    return (sa);
};


fnExcelReport();


---------------@@@@EXTRACT COLLECTION WITHOUT FOILS@@@@-------------------------------------------------------------


 function fnExcelReport() {
    var tab_text = "<table border='2px'><tr bgcolor='#87AFC6'>";
    var textRange;
    var j = 0;
    tabs = document.getElementsByClassName('CollectionTable tablesorter'); 
    tab = tabs[0];
    for (j = 0; j < tab.rows.length; j++) {
        rowInnerHtml = tab.rows[j].innerHTML;

        flag = rowInnerHtml.indexOf("Foil") === -1;
        console.log(flag);
        if (flag) {
        	tab_text = tab_text + rowInnerHtml + "</tr>";
        }
    }

    tab_text = tab_text + "</table>";
    tab_text = tab_text.replace(/<A[^>]*>|<\/A>/g, "");
    tab_text = tab_text.replace(/<img[^>]*>/gi, "");
    tab_text = tab_text.replace(/<input[^>]*>|<\/input>/gi, "");

	sa = window.open('data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,' + encodeURIComponent(tab_text));

    return (sa);
};


fnExcelReport();

------------------------------------------@@@@@@@EXTRACT ONLY FOILS@@@@----------------------------------------------------------------------


 function fnExcelReport() {
    var tab_text = "<table border='2px'><tr bgcolor='#87AFC6'>";
    var textRange;
    var j = 0;
    tabs = document.getElementsByClassName('CollectionTable tablesorter'); 
    tab = tabs[0];
    for (j = 0; j < tab.rows.length; j++) {
        rowInnerHtml = tab.rows[j].innerHTML;

        flag = rowInnerHtml.indexOf("Foil") !== -1;
        console.log(flag);
        if (flag) {
        	tab_text = tab_text + rowInnerHtml + "</tr>";
        }
    }

    tab_text = tab_text + "</table>";
    tab_text = tab_text.replace(/<A[^>]*>|<\/A>/g, "");
    tab_text = tab_text.replace(/<img[^>]*>/gi, "");
    tab_text = tab_text.replace(/<input[^>]*>|<\/input>/gi, "");

	sa = window.open('data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,' + encodeURIComponent(tab_text));

    return (sa);
};


fnExcelReport();