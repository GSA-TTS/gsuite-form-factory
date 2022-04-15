// 20220413
// v 0.0.1
const FORM_ID_INDEX = 5;
const FORM_ID_LOCATION = "A1:B1";
const FORM_TITLE_LOCATION = "B2";
const FORM_PROMPT_LOCATION = "B3";
const FORM_DATE_OPTION_RANGE = "B4:B20";

function onFormSubmit(e) {
  // Event reference
  // https://developers.google.com/apps-script/guides/triggers/events?hl=en
  // The event reference handles all kinds of good things. Amongst them are
  // the response object that was just generated (e.response) and an object
  // representing the form itself (e.source). The latter is equivalent to what
  // would be returned by FormsApp if you were opening the source form 
  // from a GDocs UID.
  var their_response = e.response;
  var form = e.source;

  // Now, I can get all of the responses on the form, and I can edit the form as well.
  // I want to get the item from the form, filter out the current choice,
  // remove the item from the form, and insert a new (shorter) set of choices.
  // Easy-peasy. (O_o)

  // The multiple-choice question MUST come last. 
  // Why MUST?
  // It's a design constraint. It makes it easy to know where to put
  // the multiple choice in the template, and where to obtain it at submission time.
  // It comes back as an interface, so it has to be cast. 
  // No, I'm not sure what kind of Javascript this is; perhaps it is Typescript?
  var the_last_item = form.getItems().reverse()[0]
  var the_mc_item = the_last_item.asMultipleChoiceItem();
  // This... should be the text of what they submitted. Responses come into the 
  // event list in form order, so the last response will be the multiple choice response.
  // I don't think we get an index; I think we get the text only. (FIXME: investigate further.)
  var the_last_response = their_response.getItemResponses().reverse()[0];
  the_current_mc_response_value = the_last_response.getResponse();

  // I'm going to walk through the choices that are on the form itself.
  // So, this takes the item from the form and grabs all the Choice objects in an array.
  current_choices = the_mc_item.getChoices();
  new_choices = [];

  // Now, walk through those choices and compare the value of each choice object 
  // (or, the text the researcher entered for each choice) and compare it to the response
  // value (the text of the choice the user selected). 
  for (var ndx = 0 ; ndx < current_choices.length ; ndx++) {
    if (current_choices[ndx].getValue() == the_current_mc_response_value) {
      // DO NOTHING if they match. 
      // https://stackoverflow.com/questions/21634886/what-is-the-javascript-convention-for-no-operation
      ()=>{}; 
    } else {
      // Keep anything that doesn't match, so we can update the form with remaining options.
      new_choices.push(current_choices[ndx]);
    }
  }

  // Update the multiple choice object in the form *in place* with the remaining choices.
  the_mc_item.setChoices(new_choices);
}

// createNewForm :: none -> none 
// Initiates a new process driver with constants for where data is located
// in the spreadsheet and does the work. This is wrapped in a function
// for calling from either the debugger or from the menu that is created onOpen.
function createNewForm() {
  d = new Driver();
  d.init({ 
      "form_id_index" : FORM_ID_INDEX,
      "form_id_location" : FORM_ID_LOCATION,
      "form_title_location" : FORM_TITLE_LOCATION,
      "form_prompt_location" : FORM_PROMPT_LOCATION,
      "form_date_option_range" : FORM_DATE_OPTION_RANGE,
    }); 
  d.createForm();
}

// onOpen :: none -> none
// Invoked when we open the factory spreadsheet.
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Form Factory')
      .addItem('Create Form', 'createNewForm')
      .addToUi();
}

// class Driver
// Encapsulates the work of reading from the spreadsheet (a very fragile kind of database)
// and builds a new GForm for the researcher.
var Driver = function () {
  // FIELDS
  this.form_id_index = null;
  this.form_id_location = null;
  this.form_title_location = null;
  this.form_prompt_location = null;
  this.form_date_option_range = null;
  this.sheet = null;

  ///////////////////// CONSTRUCTOR, SORTA /////////////////////
  this.init = function (params) {
    this.form_id_index = params["form_id_index"];
    this.form_id_location = params["form_id_location"];
    this.form_title_location = params["form_title_location"];
    this.form_prompt_location = params["form_prompt_location"];
    this.form_date_option_range = params["form_date_option_range"];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    this.sheet = ss.getSheets()[0];
  }

  // createForm :: none -> none
  // Invoked from the form factory menu.
  // Walks the sheet and creates a new form.
  this.createForm = function() {
    id = this.getFormId(this.form_id_location);
    new_form_id = this.copyTemplate(id);
    extended = this.addDateSelection(new_form_id)
    this.addTrigger(new_form_id);
  }

  ///////////////////// METHODS /////////////////////
  // getFormId :: string -> string
  // Returns the form ID from a full URL of a Google Sheet.
  this.getFormId = function(location) {
    var range = this.sheet.getRange(location);
    // The cell reference is relative to the *range* 
    // that was just selected. Indexing starts at 1.
    // So, [1, 1] is A1 (in this case).
    var url = range.getCell(1,2).getValue();
    // In a GForms URL, the fifth element is the document Id.
    // No idea how to validate this is a valid id, other than (possibly)
    // it's really long.
    var formId = url.split("/")[this.form_id_index];
    return formId
  }

  // getMultipleChoiceOptions :: range -> arrayOf string
  // Gets an array of string values out of the sheet.
  this.getMultipleChoiceOptions = function() {
    var sheet_values = this.sheet.getRange(this.form_date_option_range).getValues();
    // Values come back as a rectangular array. I want a column, 
    // but it will therefore be a list of lists, where each sublist is 
    // of length one. Hence the indexing and 'if' statements in the loop.
    var values = []
    for (var ndx = 1 ; ndx < sheet_values.length; ndx++) {
      if (sheet_values[ndx].length != 0) {
        var v = sheet_values[ndx][0]
        // Try and grab only values that were entered. Don't preserve empty cells.
        if ((typeof v === 'string' || v instanceof String) && (v.length > 0)) {
          values.push(v);
        }
      }
    }
    return values;
  }

  // copyTemplate :: string -> form_id
  // Takes a string representing a GForm Id, and returns
  // a reference a new form object representing a copy. The copy
  // is made in the same location as the template.
  this.copyTemplate = function(formId) {
    var the_template = DriveApp.getFileById(formId);
    var parents = the_template.getParents();
    var immediate = parents.next();
    var new_form_name = this.sheet.getRange(this.form_title_location).getCell(1,1).getValue()
    var new_form = the_template.makeCopy(new_form_name, immediate);
    var new_form_id = new_form.getId();
    return new_form_id
  }

  // addDateSelection :: form -> form
  // Takes a form and adds a multiple choice selection as 
  // defined in the driver spreadsheet. Returns the same
  // form for... chaining? 
  this.addDateSelection = function(form_id) {
    var form = FormApp.openById(form_id);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    var prompt = sheet.getRange(this.form_prompt_location).getCell(1, 1).getValue();
    var choices_text = this.getMultipleChoiceOptions()
    var choice_objects = [];

    var item = form.addMultipleChoiceItem();
    for (var ndx = 0 ; ndx < choices_text.length ; ndx++) {
      choice_objects.push(item.createChoice(choices_text[ndx]));
    }
    item.setTitle(prompt);
    item.setChoices(choice_objects);

  }

  // Adds a new trigger to the form.
  this.addTrigger = function(form_id) {
    ScriptApp.newTrigger('onFormSubmit')
      .forForm(form_id)
      .onFormSubmit()
      .create();
  }
} // End Driver class
