#target indesign;
#targetengine "session";
//////////////

//Resetando as preferências de PDF do Indesign
app.pdfExportPreferences.pageRange = PageRange.ALL_PAGES;
app.pdfExportPreferences.viewPDF = false;

if (app.documents.length > 0) {
  // Algumas variáveis globais
  var doc = app.activeDocument;
  var docPath = doc.filePath;
  var docname = doc.name.split(".")[0];
  var xmlRoot = doc.xmlElements[0];
  var selected_json;
  var selected_folder;
  var myPresets = app.pdfExportPresets.everyItem().name;
  var jsonLength = 0;
  var temPresetPdf = 0;
  var temFolderBase = 0;
  var temJson = 0;

  // Salvando uma cópia do arquivo para manter o original intacto.
  doc.save(new File(docGerando));
  var docGerando = docPath + "/" + docname + "_gerando.indd";
  doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.PIXELS;
  doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.PIXELS;

  //ANCHOR INTERFACE PRINCIPAL

  // Alguns parâmetros para serem usados em toda a interface
  var statictext_size = [0, 0, 460, 20];
  var edittext_size = [0, 0, 460, 30];

  // Cria a caixa de diálogo inicial
  var main_window = new Window("palette", "Gerador de arquivos variáveis");
  main_window.margins = [15, 20, 15, 20];

  var main_group = main_window.add("group", undefined, "");
  main_group.orientation = "row";
  main_group.alignChildren = "top";
  main_group.maximumSize.height = 600;

  var panel_group = main_group.add("group", undefined, "");
  panel_group.orientation = "column";
  panel_group.alignment = "fill";
  panel_group.alignChildren = "left";
  panel_group.maximumSize.height = 5000;
  panel_group.spacing = 40;

  var main_scrollbar = main_group.add("scrollbar");
  main_scrollbar.stepdelta = main_scrollbar.jumpdelta = 15;
  main_scrollbar.preferredSize.width = 12;
  main_scrollbar.preferredSize.height = 600;
  main_scrollbar.maxvalue = 1100;

  main_scrollbar.onChanging = function () {
    panel_group.location.y = -1 * this.value;
  };

  panel_group.addEventListener("scroll", function () {
    panel_group.location.y = -1 * this.value;
  });

  // Painel para seleção de preset de PDF
  var pdf_preset_panel = panel_group.add(
    "panel",
    undefined,
    "Selecione o preset do PDF"
  );
  pdf_preset_panel.margins = 20;
  pdf_preset_panel.minimumSize.width = 500;
  var pdf_preset_panel_drop01 = pdf_preset_panel.add(
    "dropdownlist",
    edittext_size,
    undefined,
    { items: myPresets }
  );

  // Painel para seleção da pasta onde os arquivos serão salvos
  var path_panel = panel_group.add(
    "panel",
    undefined,
    "Informe a pasta onde os arquivos serão salvos"
  );
  path_panel.alignChildren = "left";
  path_panel.margins = 20;
  path_panel.minimumSize.width = 500;
  var path_panel_txt01 = path_panel.add("edittext", edittext_size, "");
  var path_panel_btn01 = path_panel.add(
    "button",
    undefined,
    "Selecionar Pasta"
  );

  //Painel para seleção do arquivo json.
  var json_panel = panel_group.add(
    "panel",
    undefined,
    "Informe o endereço do arquivo json"
  );
  json_panel.alignChildren = "left";
  json_panel.margins = 20;
  json_panel.minimumSize.width = 500;
  var json_panel_txt01 = json_panel.add("edittext", edittext_size, "");
  var json_panel_btn01 = json_panel.add("button", undefined, "Selecionar Json");

//Painel para configuração dos campos no documento JSONMerge
var merge_panel = panel_group.add(
    "panel",
    undefined,
    "Marque no documento os campos a serem mesclados."
  );
  merge_panel.alignChildren = "left";
  merge_panel.margins = 20;
  merge_panel.minimumSize.width = 500;
  merge_panel.orientation = "row";
  var listbox_merge = merge_panel.add("listbox", undefined);
  listbox_merge.preferredSize = [180, 200];

  //Painel para configuração dos nomes das pastas
  var folder_panel = panel_group.add(
    "panel",
    undefined,
    "Informe para quais níveis (em ordem) serão criadas pastas"
  );
  folder_panel.alignChildren = "left";
  folder_panel.margins = 20;
  folder_panel.minimumSize.width = 500;
  folder_panel.orientation = "row";
  var listbox_folder_src = folder_panel.add("listbox", undefined);
  listbox_folder_src.preferredSize = [180, 200];
  var group_btn_move_folder = folder_panel.add("group");
  group_btn_move_folder.orientation = "column";
  var button_add_folder = group_btn_move_folder.add("button", undefined, "→");
  var button_remove_folder = group_btn_move_folder.add(
    "button",
    undefined,
    "←"
  );
  var listbox_folder_target = folder_panel.add("listbox", undefined, "");
  listbox_folder_target.preferredSize = [180, 200];

  //Painel para configuração dos nomes dos arquivos
  var file_panel = panel_group.add("panel", undefined, "Nome do arquivo pdf");
  file_panel.alignChildren = "left";
  file_panel.margins = 20;
  file_panel.preferredSize.width = 500;
  var prefix_group = file_panel.add("group");
  prefix_group.orientation = "column";
  prefix_group.alignChildren = "left";

  prefix_group.margins = [0, 0, 0, 20];
  var file_panel_prefix_check = prefix_group.add(
    "checkbox",
    undefined,
    "Usar prefixo no nome arquivo"
  );
  var file_panel_prefix_label = prefix_group.add(
    "statictext",
    undefined,
    "Prefixo"
  );
  var file_panel_prefix_txt = prefix_group.add("edittext", edittext_size, "");
  file_panel_prefix_txt.margins = [0, 0, 0, 20];

  var statictext_file_src = file_panel.add(
    "statictext",
    undefined,
    "Selecione as variáveis (em ordem) que irão compor o nome do arquivo"
  );

  var list_box_file_group = file_panel.add("group", undefined);
  list_box_file_group.orientation = "row";
  list_box_file_group.margins = [0, 0, 0, 20];

  var listbox_file_src = list_box_file_group.add("listbox", undefined);
  listbox_file_src.preferredSize = [180, 200];

  var group_btn_move_file = list_box_file_group.add("group");
  group_btn_move_file.orientation = "column";
  var button_add_file = group_btn_move_file.add("button", undefined, "→");
  var button_remove_file = group_btn_move_file.add("button", undefined, "←");
  var listbox_file_target = list_box_file_group.add("listbox", undefined, "");
  listbox_file_target.preferredSize = [180, 200];

  var file_panel_sufix_check = file_panel.add(
    "checkbox",
    undefined,
    "Usar sufixo no nome do arquivo"
  );
  var file_panel_sufix_label = file_panel.add(
    "statictext",
    undefined,
    "Sufixo"
  );
  var file_panel_sufix_txt = file_panel.add("edittext", edittext_size, "");

  var static_file_counter = panel_group.add(
    "statictext",
    undefined,
    "Aguardando importação do json para calcular o número de arquivos a serem gerados."
  );

  var btn_group = panel_group.add("group");
  btn_group.minimumSize.width = 500;
  btn_group.alignment = "fill fill";

  var main_ok_btn = btn_group.add("button", undefined, "OK");
  main_ok_btn.enabled = false;
  var main_cancel_btn = btn_group.add("button", undefined, "Cancelar");

  //--------  FUNÇÕES DOS BOTÕES DA INTERFACE NA ORDEM EM QUE FORAM CRIADOS  --------//

  //ANCHOR - Função checa preset do pdf
  pdf_preset_panel_drop01.onChange = function () {
    temPresetPdf = 1;

    checkFields();
  };

  //ANCHOR - Função que seleciona pasta para salvamento
  path_panel_btn01.onClick = function () {
    selected_folder = Folder.selectDialog("Selecione uma pasta").absoluteURI;
    path_panel_txt01.text = selected_folder;
    var logFile = new File(selected_folder + "/log.txt");
    logFile.open("w");
    logFile.write(
      "Gerador de arquivos variáveis iniciado em: " +
        new Date() +
        "\n\n ----------------------------------------------------\n\n"
    );
    logFile.close();
    temFolderBase = 1;
    checkFields();
  };

  //ANCHOR Função que seleciona JSON
  json_panel_btn01.onClick = function () {
    selected_json = File.openDialog("Selecione um arquivo JSON");
    json_panel_txt01.text = selected_json.absoluteURI;
    var keys = getKeys(selected_json)[0];
    var jsonLength = getKeys(selected_json)[1];
    static_file_counter.text = "Serão gerados " + jsonLength + " arquivos.";

    for (var i = 0; i < keys.length; i++) {
      listbox_folder_src.add("item", keys[i]);
      listbox_file_src.add("item", keys[i]);
      var list_item = listbox_merge.add("item", keys[i]);
      createTags(keys[i]);
    }
    temJson = 1;
    checkFields();
  };

  //ANCHOR Listbox_merge Funções disparadas a partir da seleção de itens no listbox do merge
  listbox_merge.onChange = function () {
    var listbox_selection = listbox_merge.selection;
    var mouse_selection = app.selection[0];
    
    if(mouse_selection != undefined) {
        xmlRoot.xmlElements.add(listbox_selection.text, mouse_selection);
    } else{
        alert("Selecione porções de texto para adicionar um campo JSON")
    }
      
  }

  // Funções disparadas a partir da seleção de itens no listbox de Folders
  listbox_folder_src.onActivate = function () {
    listbox_folder_src.onChange = function () {
      button_add_folder.enabled = true;
    };
    button_remove_folder.enabled = false;
  };

  listbox_folder_target.onActivate = function () {
    listbox_folder_target.onChange = function () {
      button_remove_folder.enabled = true;
    };
    button_add_folder.enabled = false;
  };

  if (listbox_folder_src.selection == null) {
    button_add_folder.enabled = false;
  }

  if (listbox_folder_target.selection == null) {
    button_remove_folder.enabled = false;
  }

  // Migrando itens da origem para o destino de folders
  button_add_folder.onClick = function () {
    listbox_folder_target.add("item", listbox_folder_src.selection);
    listbox_folder_src.remove(listbox_folder_src.selection);
    button_add_folder.enabled = false;
  };
  // Retornando itens do destino para a origem de folders
  button_remove_folder.onClick = function () {
    listbox_folder_src.add("item", listbox_folder_target.selection);
    listbox_folder_target.remove(listbox_folder_target.selection);
    button_remove_folder.enabled = false;
  };

  // Funções disparadas a partir da seleção de itens no listbox de Files
  listbox_file_src.onActivate = function () {
    listbox_file_src.onChange = function () {
      button_add_file.enabled = true;
    };
    button_remove_file.enabled = false;
  };

  listbox_file_target.onActivate = function () {
    listbox_file_target.onChange = function () {
      button_remove_file.enabled = true;
    };
    button_add_file.enabled = false;
  };

  if (listbox_file_src.selection == null) {
    button_add_file.enabled = false;
  }

  if (listbox_file_target.selection == null) {
    button_remove_file.enabled = false;
  }

  // Migrando itens da origem para o destino de files
  button_add_file.onClick = function () {
    listbox_file_target.add("item", listbox_file_src.selection);
    listbox_file_src.remove(listbox_file_src.selection);
    button_add_file.enabled = false;
    checkFields();
  };
  // Retornando itens do destino para a origem de files
  button_remove_file.onClick = function () {
    listbox_file_src.add("item", listbox_file_target.selection);
    listbox_file_target.remove(listbox_file_target.selection);
    button_remove_file.enabled = false;
    checkFields();
  };

  //ANCHOR Funções para habilitar e desabilitar o campo de prefixo com base no edittext de prefixo
  file_panel_prefix_txt.onChanging = function () {
    if (file_panel_prefix_txt.text == "") {
      file_panel_prefix_check.value = false;
    } else {
      file_panel_prefix_check.value = true;
    }
  };

  //ANCHOR Funções para habilitar e desabilitar o campo de sufixo com base no edittext de sufixo
  file_panel_sufix_txt.onChanging = function () {
    if (file_panel_sufix_txt.text == "") {
      file_panel_sufix_check.value = false;
    } else {
      file_panel_sufix_check.value = true;
    }
  };

  //ANCHOR Função do botão OK - Aciona a função principal
  main_ok_btn.onClick = function () {
    main_window.close();
    main(selected_json, selected_folder);
  };

  //ANCHOR Função do botão Cancelar - Encerra o script
  main_cancel_btn.onClick = function () {
    main_window.close();
  };

  //ANCHOR Exibindo a interface
  var show_panel = main_window.show();

  //ANCHOR FUNÇÕES GERAIS

  //ANCHOR Função principal
  function main(selected_json, selected_folder) {
    var logFile = File(selected_folder + "/log.txt");
    // Abrindo o arquivo JSON, buscando as chaves e convertendo os campos do datamerge em variables
    selected_json.open("r");
    var readJSON = selected_json.read();
    var objetoJSON = eval("(" + readJSON + ")");
    var jsonItens = objetoJSON.itens;
    jsonLength = jsonItens.length;

    // Criando a barra de progresso com base no tamanho do json
    var progress = janelaProgresso.add("progressbar", undefined, 0, jsonLength);
    janelaProgresso.show();
    progress.value = 0;
    progress.size = [480, 20];

    var keys = getKeys(selected_json)[0];

    // Percorrendo item por item do json para gerar os arquivos.
    // 1 . Mesclamos o conteúdo com o valor de cada chave de cada item.
    // 2 . Configuramos a pasta onde o arquivo será criado.
    // 3 . Criamos o arquivo pdf, slavando na pasta criada.
    for (i = 0; i < jsonItens.length; i++) {
      progress.value = i + 1;
      janelaProgressoContador.text =
        "Gerando o arquivo " + (i + 1) + "/" + jsonLength;
      var jsonItem = jsonItens[i];
      
      for (j = 0; j < keys.length; j++) {
        jsonKey = keys[j];

        textTag = xmlRoot.evaluateXPathExpression("//"+ jsonKey);
        alert(jsonKey + " - " + textTag);
        if(textTag.length > 0){

    for (var k = 0; k < textTag.length; k++) {
        const element = textTag[k];

        element.xmlContent.contents = jsonItem[jsonKey].toString()
        
    }
    
}
      }

      var path = selected_folder;
      var pdfname;
      setPath();
      setFileName();

      /* var tipo_pdf = app.pdfExportPresets.item(
        String(pdf_preset_panel_drop01.selection)
      );
      doc.exportFile(ExportFormat.PDF_TYPE, new File(pdfname), false, tipo_pdf);

      logFile.open("a");
      logFile.write(pdfname + " gerado com sucesso.\n");
      logFile.close(); */

      function setPath() {
        // Percorre o listbox de folders com a seleção do usuário para criar as pastas
        for (var i = 0; i < listbox_folder_target.items.length; i++) {
          var folder = jsonItem[listbox_folder_target.items[i].text].toString();
          path = path + "/" + folder;
          new Folder(path).create();
        }

        // Seta o caminho do arquivo
        pdfname = path + "/";
      }

      // Percorre o listbox de files com a seleção do usuário para setar as variáveis e incrementa o nome do arquivo
      function setFileName() {
        // Checa se há prefixo e incrementa o nome do arquivo
        if (file_panel_prefix_check.value == true) {
          pdfname += file_panel_prefix_txt.text + "_";
        }

        for (var i = 0; i < listbox_file_target.items.length; i++) {
          var listItem = jsonItem[listbox_file_target.items[i].text].toString();
          pdfname = pdfname + listItem + "_";
        }

        // Checa se há sufixo e incrementa o nome do arquivo
        if (file_panel_sufix_check.value == true) {
          pdfname += "_" + file_panel_sufix_txt.text;
        }

        // Adiciona extensão .pdf e conlui a configuração do nome do arquivo
        pdfname += ".pdf";

        // Remove underlines excessivos
        pdfname = pdfname.replace("__", "_");

        // Remove underlines excessivos
        pdfname = pdfname.replace("_.pdf", ".pdf");
      }
    }

    janelaProgressoContador.text =
      jsonLength + " arquivos gerados com sucesso!";
  }

  // Função que busca chaves do json para preencher os listboxs da interface
  function getKeys(jsonFie) {
    jsonFie.open("r");
    var readJSON = jsonFie.read();
    var objetoJSON = eval("(" + readJSON + ")");
    var jsonKeys = objetoJSON.itens[0];
    var jsonLength = objetoJSON.itens.length;

    var getJsonKeys = function (jsonObject) {
      var arrayWithKeys = [],
        jsonObject;
      for (key in jsonObject) {
        if (
          key !== undefined &&
          key !== "toJSONString" &&
          key !== "parseJSON"
        ) {
          arrayWithKeys.push(key);
        }
      }
      return arrayWithKeys;
    };

    var keys = getJsonKeys(jsonKeys);
    return [keys, jsonLength];
  }

  // Função que checa se um preset de pdf foi selecionado
  function checkFields() {
    if (
      temPresetPdf == 1 &&
      temFolderBase == 1 &&
      temJson == 1 &&
      listbox_file_target.items.length > 0
    ) {
      main_ok_btn.enabled = true;
    } else {
      main_ok_btn.enabled = false;
    }
  }

  function createTags(tag) {
    doc.xmlElements[0].xmlElements.add(tag);
  }

  // --------- INTERFACE DO PROGRESSO --------- //
  var janelaProgresso = new Window("palette");
  janelaProgresso.text = "Progresso";
  janelaProgresso.margins = [10, 10, 10, 10];
  janelaProgresso.alignChildren = "left";
  janelaProgresso.preferredSize.width = 500;
  var janelaProgressoContador = janelaProgresso.add(
    "statictext",
    undefined,
    "Gerando o arquivo 0/" + jsonLength
  );
  janelaProgressoContador.preferredSize.width = 480;
} else {
  alert(
    "Nenhum documento aberto! Abra o arquivo para geração dos arquivos antes de continuar."
  );
}
