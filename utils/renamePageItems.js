var doc = app.activeDocument;
var pageItems = doc.pageItems;

for (var i = 0; i < pageItems.length; i++) {
    pageItemName = pageItems[i].name;
    
    if(pageItemName.match('_ESC_')){
        var newName = pageItemName.replace('_ESC_', '_REG_');
        pageItems[i].name = newName;
    }
    
    /* else if(pageItemName.match('_lp_')){
        var newName = pageItemName.replace('_lp_', '_escrita_');
        pageItems[i].name = newName;
    } */

    /* else if(pageItemName.match('_1b')){
        var newName = pageItemName.replace('_1b', '_3b');
        pageItems[i].name = newName;
    } */
}