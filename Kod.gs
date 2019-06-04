var source = SpreadsheetApp.getActiveSpreadsheet();

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Drukuj')
  .addItem('Karty NIE ODZIWEDZAĆ', 'dialog')
  .addToUi();
}

function test() {
    generate_cards(false, 300, 500, false) 
}


function generate_cards(only_with_content, range_from, range_to, empty_template) {
  var new_sheet = source.getSheetByName("NIE ODWIEDZAĆ")
  var all_data =  new_sheet.getDataRange()
  var rows_count = all_data.getLastRow()
  var all_deny_data = new_sheet.getRange("A3:C"+rows_count).getValues()
  var cards = HtmlService.createTemplateFromFile("deny_cards.tpl")
  var data_to_display = []
  var deny_data = []

  if (empty_template) {
      for (var h=range_from; h <= range_to; h++) {
          deny_data.push([h, "", ""])
      }
      data_to_display = deny_data
  } else {
    
    
      if (range_from && range_to) {
          for (var k=0; k < all_deny_data.length; k++) {
              var terr_nr =  all_deny_data[k][0]
              if (terr_nr >= range_from && terr_nr <= range_to)
                  deny_data.push(all_deny_data[k])
          }
      } else {
          deny_data = all_deny_data
      }

      if (! only_with_content) {
          if (deny_data.length == 0)
            throw "Błędny zakres. Nie znalezniono pasujących danych w zakresie."
          
          var full_deny_data = []
          var previous_terr_nr = deny_data[0][0];
          
  
          for (var i=0; i < deny_data.length; i++) {
              var terr_nr =  deny_data[i][0]
              
              for (j=1; j < terr_nr - previous_terr_nr; j++) {
                var tmp_array = [ previous_terr_nr + j, "", ""]
                full_deny_data.push(tmp_array)
              }
              full_deny_data.push(deny_data[i])
              previous_terr_nr = terr_nr
          }
          data_to_display = full_deny_data
      } else {
      
      if (deny_data.length == 0)
          throw "Brak danych dla podanego zakresu."
        data_to_display = deny_data
      }
  }
  cards.deny_data = data_to_display
  cards.rows = rows_count
  return cards.evaluate().getContent()
}

function dialog() {
  var active_sheet_name = source.getActiveSheet().getSheetName()
  var html = HtmlService.createTemplateFromFile("dialog.tpl")
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(400).setHeight(335), 'Drukuj ' + active_sheet_name)
}

function include(File) {
  return HtmlService.createHtmlOutputFromFile(File).getContent();
};