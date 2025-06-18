var main = function () {
  var pages = app.documents[0].spreads.everyItem().getElements();
  var layerName = "redimensiona";
  var layerArr, apis;
  for (var i = 0; i < pages.length; i++) {
    layerArr = new Array();
    apis = pages[i].pageItems;
    for (var j = 0; j < apis.length; j++) {
      if (apis[j].itemLayer.name == layerName) {
        layerArr.push(apis[j]);
      }
    }

    layerArr.sort(function (a, b) {
      return a.name.localeCompare(b.name);
    });

    for (var k = 0; k < layerArr.length; k++) {
      try {
        layerArr[k].sendToBack();
      } catch (e) {}
    }
  }
};

main();
