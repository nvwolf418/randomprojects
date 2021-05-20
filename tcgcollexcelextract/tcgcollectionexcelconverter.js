/*
*Steps:
* 1) Go to the tcgplayer collect link and sort it
* 2) Copy everything between dotted lines below
* 3) select address bar where tcgplayer url is and press ctrl + a, then type- 
*     javascript:
* 4) press ctrl + v
* 5) press enter, download should start!
*/



-------------------------------------------------------------------------


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
        if (flag)
            tab_text = tab_text + rowInnerHtml + "</tr>";
    }

    tab_text = tab_text + "</table>";
    tab_text = tab_text.replace(/<A[^>]*>|<\/A>/g, "");
    tab_text = tab_text.replace(/<img[^>]*>/gi, "");
    tab_text = tab_text.replace(/<input[^>]*>|<\/input>/gi, "");

    var ua = window.navigator.userAgent;
    var msie = ua.indexOf("MSIE ");

    if (msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./))
    {
        txtArea1.document.open("txt/html", "replace");
        txtArea1.document.write(tab_text);
        txtArea1.document.close();
        txtArea1.focus();
        sa = txtArea1.document.execCommand("SaveAs", true, "Say Thanks to Sumit.xls");
    } else
        sa = window.open('data:application/vnd.ms-excel,' + encodeURIComponent(tab_text));

    return (sa);
};


fnExcelReport();


----------------------------------------------------------------------------