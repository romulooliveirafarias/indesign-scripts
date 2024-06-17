#target indesign;
#targetengine "session";
//////////////

//Resetando as preferências de PDF do Indesign
app.pdfExportPreferences.pageRange = PageRange.ALL_PAGES;
app.pdfExportPreferences.viewPDF = false;

var doc = app.activeDocument;
var docPath = doc.filePath;
var docname = doc.name.split(".")[0];
var selected_json;
var selected_folder;
var myPresets = app.pdfExportPresets.everyItem().name;
myPresets.unshift("- Select Preset -");
/* var docGerando = docPath + "/" + docname + "_gerando.indd"; */

var statictext_size = [0, 0, 200, 20]
var edittext_size = [0, 0, 380, 30]
/* doc.save(new File(docGerando)); */
doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.PIXELS;
doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.PIXELS;

// Cria a caixa de diálogo inicial
var main_window = new Window('palette', "Gerador de arquivos variáveis");
main_window.margins = [15, 20, 15, 20];

var main_group = main_window.add("group", undefined, "");
main_group.orientation = "row";
main_group.alignChildren = "top";
main_group.maximumSize.height = 520;

var panel_group = main_group.add("group", undefined, "");
panel_group.orientation = "column";
panel_group.alignment = "fill";
panel_group.alignChildren = "left";
panel_group.maximumSize.height = 5000;
panel_group.spacing = 40;

var main_scrollbar = main_group.add("scrollbar");
main_scrollbar.stepdelta = main_scrollbar.jumpdelta = 15;
main_scrollbar.preferredSize.width = 16;
main_scrollbar.preferredSize.height = 500;
main_scrollbar.maxvalue = 750;

main_scrollbar.onChanging = function () {
  panel_group.location.y = -1 * this.value;
  }

var pdf_preset_panel = panel_group.add("panel", undefined, "Selecione o preset do PDF");
var pdf_preset_panel_drop01 = pdf_preset_panel.add('dropdownlist',edittext_size,undefined,{items:myPresets});

var path_panel = panel_group.add("panel", undefined, "Informe a pasta onde os arquivos serão salvos");
path_panel.alignChildren = "left";
path_panel.margins.top = 20;
path_panel.preferredSize.width = 400;
var path_panel_txt01 = path_panel.add("edittext", edittext_size, "");
var path_panel_btn01 = path_panel.add("button", undefined, "Selecionar Pasta");
path_panel_btn01.onClick = function () {
  selected_folder = Folder.selectDialog("Selecione uma pasta").absoluteURI + "/";;
  path_panel_txt01.text = selected_folder;
  filename = selected_folder;
}

var json_panel = panel_group.add("panel", undefined, "Informe o endereço do arquivo json");
json_panel.alignChildren = "left";
json_panel.margins.top = 20;
json_panel.preferredSize.width = 400;
var json_panel_txt01 = json_panel.add("edittext", edittext_size, "");
var json_panel_btn01 = json_panel.add("button", undefined, "Selecionar Json");
json_panel_btn01.onClick = function () {
  selected_json = File.openDialog("Selecione um arquivo JSON");
  json_panel_txt01.text = selected_json.absoluteURI;
}

var folder_panel = panel_group.add("panel", undefined, "Para quais níveis serão criadas pastas?");
folder_panel.alignChildren = "left";
folder_panel.margins.top = 20;
folder_panel.preferredSize.width = 400;
var folder_panel_check01 = folder_panel.add("checkbox", undefined, "Rede");
var folder_panel_label01 = folder_panel.add("statictext", undefined, "Variável da Rede");
var folder_panel_txt01 = folder_panel.add("edittext", edittext_size, "");
var folder_panel_check02 = folder_panel.add("checkbox", undefined, "Regional");
var folder_panel_label02 = folder_panel.add("statictext", undefined, "Variável da Regional");
var folder_panel_txt02 = folder_panel.add("edittext", edittext_size, "");
var folder_panel_check03 = folder_panel.add("checkbox", undefined, "Município");
var folder_panel_label03 = folder_panel.add("statictext", undefined, "Variável do Município");
var folder_panel_txt03 = folder_panel.add("edittext", edittext_size, "");

var file_panel = panel_group.add("panel", undefined, "Nome do arquivo pdf");
file_panel.alignChildren = "left";
file_panel.margins.top = 20;
file_panel.preferredSize.width = 400;
var file_panel_check01 = file_panel.add("checkbox", undefined, "Usar Código da escola");
var file_panel_label01 = file_panel.add("statictext", undefined, "Variável do código da escola");
var file_panel_txt01 = file_panel.add("edittext", edittext_size, "");
var file_panel_check02 = file_panel.add("checkbox", undefined, "Usar nome da escola");
var file_panel_label02 = file_panel.add("statictext", undefined, "Variável do nome da escola");
var file_panel_txt02 = file_panel.add("edittext", edittext_size, "");
var file_panel_check03 = file_panel.add("checkbox", undefined, "Usar prefixo no nome arquivo");
var file_panel_label03 = file_panel.add("statictext", undefined, "Prefixo");
var file_panel_txt03 = file_panel.add("edittext", edittext_size, "");
var file_panel_check04 = file_panel.add("checkbox", undefined, "Usar sufixo no nome do arquivo");
var file_panel_label04 = file_panel.add("statictext", undefined, "Sufixo");
var file_panel_txt04 = file_panel.add("edittext", edittext_size, "");

var btn_group = panel_group.add("group");
btn_group.minimumSize.width = 400;
btn_group.alignment = "fill fill";
var main_ok_btn = btn_group.add("button", undefined, "OK");
var main_cancel_btn = btn_group.add("button", undefined, "Cancelar");

main_ok_btn.onClick = function () {
  main(selected_json, selected_folder);
}

main_cancel_btn.onClick = function () {
  main_window.close();
}

function main(selected_json, selected_folder) {
main_window.close();
selected_json.open("r");
var readJSON = selected_json.read();
var objetoJSON = eval("(" + readJSON + ")");
var jsonItens = objetoJSON.itens;
var jsonKeys = objetoJSON.itens[0];

var getJsonKeys = function (jsonObject) {
  var arrayWithKeys = [],
    jsonObject;
  for (key in jsonObject) {
    if (key !== undefined && key !== "toJSONString" && key !== "parseJSON") {
      arrayWithKeys.push(key);
    }
  }
  return arrayWithKeys;
};

var keys = getJsonKeys(jsonKeys);

app.changeTextPreferences.changeTo = "";

for (j = 0; j < keys.length; j++) {
  jsonKey = keys[j];

  app.findTextPreferences = app.changeTextPreferences = null;
  var buscar = (app.findTextPreferences.findWhat = "<<" + jsonKey + ">>");

  var allFounds = doc.findText();

  for (var i = 0; i < allFounds.length; i++) {
    var curFound = allFounds[i];

    curFound.select();

    var newVariable = doc.textVariables.add();
    newVariable.name = jsonKey;
    newVariable.variableType = VariableTypes.CUSTOM_TEXT_TYPE;
    newVariable.variableOptions.contents = jsonKey;
    app.selection[0].textVariableInstances.add({
      associatedTextVariable: newVariable,
    });
    app.changeTextPreferences.changeTo = "";
    doc.changeText();
  }
}

var docVariables = doc.textVariables;

for (i = 0; i < jsonItens.length; i++) {
  var jsonItem = jsonItens[i];

  for (k = 0; k < docVariables.length; k++) {
    var docVariable = docVariables[k];

    for (j = 0; j < keys.length; j++) {
      jsonKey = keys[j];

      if (docVariable.name == keys[j]) {
        docVariable.variableOptions.contents = jsonItem[jsonKey].toString();
      }
    }
  }
  createPath();

  function createPath(){
    var rede;
    var regional;
    var municipio;
    var CDescola;
    var escola;
    
    if (folder_panel_check01.value == true) {  
    rede = jsonItem[folder_panel_txt01.text].toString();
    pathRede = selected_folder + "/" + rede;
    pastaRede = new Folder(pathRede).create();
  } else{
    pathRede  = selected_folder;
  }
  
  if (folder_panel_check02.value == true) {
    regional = jsonItem[folder_panel_txt02.text].toString();
    pathRegional = pathRede + "/" + regional;
    pastaRegional = new Folder(pathRegional).create();
  } else{
    pathRegional  = pathRede;
  }

  if (folder_panel_check03.value == true) {
    municipio = jsonItem[folder_panel_txt03.text].toString();
    pathMunicipio = pathRegional + "/" + municipio;
    pastaMunicipio = new Folder(pathMunicipio).create();
  } else{
    pathMunicipio  = pathRegional;
  }

  var pdfname = pathMunicipio + "/";

if (file_panel_check03.value == true) {
    pdfname += file_panel_txt03.text;
}
if (file_panel_check02.value == true) {
    pdfname += "_" + jsonItem[file_panel_txt02.text].toString();
}
if (file_panel_check01.value == true) {
    pdfname += "_" + jsonItem[file_panel_txt01.text].toString();
}
if (file_panel_check04.value == true) {
    pdfname += "_" + file_panel_txt04.text;
}
pdfname += ".pdf";

  var tipo_pdf = app.pdfExportPresets.item(String(pdf_preset_panel_drop01.selection));
  doc.exportFile(ExportFormat.PDF_TYPE, new File(pdfname), false, tipo_pdf);
}
}

}

var show_panel = main_window.show();
