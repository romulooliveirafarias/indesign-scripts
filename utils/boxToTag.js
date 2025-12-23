if (app.documents.length === 0) {
  alert("Nenhum documento aberto.");
  exit();
}

var doc = app.activeDocument;

// Garante que exista um elemento raiz
var root;
if (doc.xmlElements.length === 0) {
  root = doc.xmlElements.add("root");
} else {
  root = doc.xmlElements[0];
}

// Função para obter ou criar a tag
function getOrCreateTag(tagName) {
  var tag = doc.xmlTags.itemByName(tagName);

  if (tag.isValid) {
    // A tag realmente existe
    return tag;
  } else {
    // A tag não existe, então cria
    return doc.xmlTags.add(tagName);
  }
}


var items = doc.pageItems;

for (var i = 0; i < items.length; i++) {
  var item = items[i];

  try {
    var tag = getOrCreateTag(item.name);

    // Cria elemento XML e associa ao item
    var xmlElement = root.xmlElements.add(tag);
    xmlElement.markup(item);

  } catch (e) {
    // Ignora itens que não podem ser marcados
  }
}

alert("Tags XML criadas a partir dos nomes dos objetos.");
