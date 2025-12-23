var doc = app.activeDocument;
var redimensionar = doc.pageItems.item("teste");
  var tamanho_maximo_barra = 100;

redimensionar.resize(
  CoordinateSpaces.INNER_COORDINATES,
  AnchorPoint.LEFT_CENTER_ANCHOR,
  ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
  [tamanho_maximo_barra, 10]
);
