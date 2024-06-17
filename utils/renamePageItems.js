var doc = app.activeDocument;
var pageItems = doc.pageItems;

for (var i = 0; i < pageItems.length; i++) {
    pageItemName = pageItems[i].name;
    
    if(pageItemName.match('_1B_REG')){
        var newName = pageItemName.replace('_1B_REG', '_REG_1B');
        pageItems[i].name = newName;
    }
}