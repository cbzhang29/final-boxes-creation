var slidesID = ""; //actual slides id
//var slidesID = "1KSRNvuKJRybgqiTNzpipfGVbNttjpcUJPEncX52qkNA";
//var word_arr = []; //sentences
//var type_arr = []; //types
//var page_arr = []; //page num

/*
Type 1: find half the length of the sentence and then subtract that amount from the middle of the slide (360) and that is the starting x pt, y = 284
Type 2: x1 = 125.5, y1 = 108; x2 = 125.5, y2 = 251
Type 3: x1 = , y1 = ; x2 = , y2 = ; x3 = , y3 = 
Type 4: x1 = , y1 = ; x2 = , y2 = ; x3 = , y3 = ; x4 = , y4 = ;
*/

var x = 0; //roughly x width of slide is 720  // type 2 x:127 y:110  pt.2 = x: 127 y:252
var y = 0; //roughly y height of slide is 405
var len = 9.34;
//var space = 10;

function getRssFeed() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("rss-feed-contents");
  if (cached != null) {
    return cached;
  }
  var result = UrlFetchApp.fetch("https://docs.google.com/spreadsheets/d/1i0kefZqcg66V8Mz7ahdVs3Fe1QKY__7E49mxdjI6A_Y/edit#gid=1030083970"); // takes 20 seconds
  var contents = result.getContentText();
  cache.put("rss-feed-contents", contents, 1500); // cache for 25 minutes
  return contents;
}

function main(){
  Logger.log("begin");
  //logProductInfo(); // three arrays: word_arr, type_arr, page_arr; should all be the same length
  var sheet = SpreadsheetApp.getActiveSheet(); //gets the sheet you are on
  var data = sheet.getDataRange().getValues(); //gets the values on the sheet
  slidesID = (data[1][4]); //assigns the slidesID to the value at E2 (the index starts at 0; so row 1 is actually 0)
  var presentation = Slides.Presentations.get(slidesID); //access the presentation based on ID
  var slides = presentation.slides; //get the slides from the presentation
  
 // getRssFeed();
  
  for (var z = 32; z < data.length; z++) { //starting from row 33
    var type = data[z][21];
    //data[z][6] is sentence, data[z][22] is page
    if(type == 1){
      type1(data[z][6].trim(), data[z][22], slides);
      Logger.log("Type 1 just ran.");
      //Utilities.sleep(20000);
    }
    else if (type == 2){
      var sents_arr = [data[z][6].trim(), data[z+1][6].trim()];
      type2(sents_arr, data[z][22], slides);
      z++;
      Logger.log("Type 2 just ran.");
     // Utilities.sleep(20000);
    }
    else if(type == 3){
      var sents_arr = [data[z][6].trim(), data[z+1][6].trim(), data[z+2][6].trim()];;
      type3(sents_arr, data[z][22], slides);
      z = z+2;
      Logger.log("Type 3 just ran");
      //Utilities.sleep(20000);
    }
    else if(type ==4){
      var sents_arr = [data[z][6].trim(), data[z+1][6].trim(), data[z+2][6].trim(), data[z+3][6].trim()];
      type4(sents_arr, data[z][22], slides);
      z = z+3;
      Logger.log("Type 4 just ran.");
    }
    else{
      //Logger.log("Blank cell");
    }
  }
}

function test_types(){
  Logger.log("begin");
  //logProductInfo(); // three arrays: word_arr, type_arr, page_arr; should all be the same length
  var sheet = SpreadsheetApp.getActiveSheet(); //gets the sheet you are on
  var data = sheet.getDataRange().getValues(); //gets the values on the sheet
  slidesID = (data[1][4]); //assigns the slidesID to the value at E2 (the index starts at 0; so row 1 is actually 0)
  var presentation = Slides.Presentations.get(slidesID); //access the presentation based on ID
  var slides = presentation.slides; //get the slides from the presentation
  
  for (var z = 1129; z < data.length; z++){
    var type = data[z][21];
    if(type ==4){
      var sents_arr = [data[z][6].trim(), data[z+1][6].trim(), data[z+2][6].trim(), data[z+3][6].trim()];
      type4(sents_arr, data[z][22], slides);
      z = z+3;
      Logger.log("Type 4 just ran.");
    }
  }
  
}

function type1(sentence, page, slides){
  y = 284; 
  var pageID = slides[page-1].objectId;
  var text = sentence;
  var size = text.length*len;
  var mid = size/2;
  x = 360-mid;
  
  var substrings = text.split(' ');

  for(var i = 0; i < substrings.length; i++){

    if(substrings[i].startsWith('¡') || substrings[i].startsWith("¿")|| substrings[i].startsWith("“") || substrings[i].startsWith('"')){
      var temp = substrings[i];
      x = x + len;
      if (temp[1] == '¡' || temp[1] == "¿"){
        x = x + len;
      }
      if (substrings[i].endsWith("!") || substrings[i].endsWith("?") ||substrings[i].endsWith(",") ||substrings[i].endsWith("”") ||substrings[i].endsWith(".") || substrings[i].endsWith('"')){
        var q = substrings[i].charAt(substrings[i].length-2); 
        var new_text = removePunctuation(substrings[i]);
        var wordlength = new_text.length*len; //gets the lengths of each word 

        var rect = addRectangle(slidesID, pageID, wordlength);
        changeRectangleColor(slidesID, rect);
        
        x = x + wordlength +len +len; //should be current x + wordlength + space + size of punctuation;
        if(q == "!" || q == "?" || q == "." || q == ","){
          x = x + len;
        }
      }
      else{
        var new_text = removePunctuation(substrings[i]);
        var wordlength = new_text.length*len; //gets the lengths of each word 
        var rect = addRectangle(slidesID, pageID, wordlength);
        changeRectangleColor(slidesID, rect);
        x = x + wordlength +len;
      }
    }
    else if (substrings[i].endsWith("?") || substrings[i].endsWith("!") || substrings[i].endsWith(",") || substrings[i].endsWith("”")||substrings[i].endsWith(".") || substrings[i].endsWith('"')){
      var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character
      var new_text = removePunctuation(substrings[i]);
      var wordlength = new_text.length*len;//gets the lengths of each word 
      var rect = addRectangle(slidesID, pageID, wordlength);
      changeRectangleColor(slidesID, rect);
      x = x + wordlength +len +len; 
      if(q == "!" || q == "?" || q == "." || q == ","){
        x = x +len;
      }
    }
    else if(substrings[i].startsWith("/") && substrings[i].length == 1){
      var wordlength = substrings[i].length*len; //gets the lengths of each word 
      x = x + wordlength +len;
    }
    else{
      var wordlength = substrings[i].length*len; //gets the lengths of each word 
      var rect = addRectangle(slidesID, pageID, wordlength);
      changeRectangleColor(slidesID, rect);
      x = x + wordlength +len;
    }
  Utilities.sleep(1000);}
}

function type2(sentences, page, slides){
  x = 125.5; y = 108;
  var pageID = slides[page-1].objectId; //get the pageID of the slide
  
  for(var w = 0; w < sentences.length; w++){ // loops through the number of sentences
    var text = sentences[w]; //two sentences
    if(w == 1){ //reset starting point
      x = 125.5;
      y = 251;
    }
    var substrings = text.split(' ');
  
    for(var i = 0; i < substrings.length; i++){
      if(substrings[i].startsWith('¡') || substrings[i].startsWith("¿")|| substrings[i].startsWith("“") || substrings[i].startsWith("/") || substrings[i].startsWith('"')){
        var temp = substrings[i];
        x = x + len;
        if (temp[1] == '¡' || temp[1] == "¿"){
          x = x + len;
        }
        if (substrings[i].endsWith("!") || substrings[i].endsWith("?") ||substrings[i].endsWith(",") ||substrings[i].endsWith("”") ||substrings[i].endsWith(".") || substrings[i].endsWith('"')){
          //var p = substrings[i].charAt(substrings[i].length-1); //p is the last character
          var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character
          //Logger.log("End char: " + p);
          var new_text = removePunctuation(substrings[i]);
          // Logger.log(new_text.length);
          var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths
          
          var rect = addRectangle(slidesID, pageID, wordlength);        
          changeRectangleColor(slidesID, rect);
          
          x = x + wordlength +len +len; //should be + space + size of punctuation;
          if(q == "!" || q == "?" || q == "." || q == ","){
            //Logger.log("2nd to last: " + q);
            x = x + len;
          }   
        }
        else{
          var new_text = removePunctuation(substrings[i]);
          //Logger.log(new_text);
          var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths
          //if type 2:
          //for(var j = 0; j<wordlength.length; j++){
          var rect = addRectangle(slidesID, pageID, wordlength);       
          changeRectangleColor(slidesID, rect);
          x = x + wordlength +len;
        }
      }
      else if (substrings[i].endsWith("?") || substrings[i].endsWith("!") || substrings[i].endsWith(",") || substrings[i].endsWith("”")||substrings[i].endsWith(".") || substrings[i].endsWith('"')){
        var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character
        //Logger.log("End char: " + p);
        var new_text = removePunctuation(substrings[i]);
        // Logger.log(new_text);
        var wordlength = new_text.length*len;//gets the lengths of each word //words is an array of the wordlengths
        //Logger.log("Wordlength: " + wordlength);
        //if type 2:
        //for(var j = 0; j<wordlength.length; j++){
        var rect = addRectangle(slidesID, pageID, wordlength);       
        changeRectangleColor(slidesID, rect);
        x = x + wordlength +len +len; //should be + 5 + size of punctuation;
        if(q == "!" || q == "?" || q == "." || q == ","){
          //Logger.log("2nd to last: " + q);
          x = x +len;
        }
      }
      else{
        var wordlength = substrings[i].length*len; //gets the lengths of each word
        var rect = addRectangle(slidesID, pageID, wordlength);      
        changeRectangleColor(slidesID, rect);
        x = x + wordlength +len;
      }
    Utilities.sleep(1000);}
  //Utilities.sleep(20000);
  }
}

function type3(sentences, page, slides){
  x = 125.5; y = 78;
  var pageID = slides[page-1].objectId; //get the pageID of the slide
  for(var w = 0; w < sentences.length; w++){ // loops through the number of sentences
    var text = sentences[w]; //two sentences
    if(w == 1){ //reset starting point
      x = 125.5;
      y = 182;
    }
    else if(w == 2){
      x = 125.5
      y = 286;
    }
    var text = sentences[w]; //whole sentence
    var substrings = text.split(' ');
    
    for(var i = 0; i < substrings.length; i++){
      
      if(substrings[i].startsWith('¡') || substrings[i].startsWith("¿")|| substrings[i].startsWith("“") || substrings[i].startsWith("/") || substrings[i].startsWith('"')){
        var temp = substrings[i];
        //Logger.log("Start char: " + temp[0]);
        x = x + len;
        if (temp[1] == '¡' || temp[1] == "¿"){
          x = x + len;
        }
        if (substrings[i].endsWith("!") || substrings[i].endsWith("?") ||substrings[i].endsWith(",") ||substrings[i].endsWith("”") ||substrings[i].endsWith(".") || substrings[i].endsWith('"')){
          //var p = substrings[i].charAt(substrings[i].length-1); //p is the last character
          var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character
          //Logger.log("End char: " + p);
          var new_text = removePunctuation(substrings[i]);
          //   Logger.log(new_text.length);
          var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths
          
          var rect = addRectangle(slidesID, pageID, wordlength);    
          changeRectangleColor(slidesID, rect);
          
          x = x + wordlength +len +len; //should be + space + size of punctuation;
          if(q == "!" || q == "?" || q == "." || q == ","){
            // Logger.log("2nd to last: " + q);
            x = x + len;
          }   
        }
        else{
          var new_text = removePunctuation(substrings[i]);
          // Logger.log(new_text);
          var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths
          //if type 2:
          //for(var j = 0; j<wordlength.length; j++){
          var rect = addRectangle(slidesID, pageID, wordlength);    
          changeRectangleColor(slidesID, rect);
          x = x + wordlength +len;
        }
      }
      else if (substrings[i].endsWith("?") || substrings[i].endsWith("!") || substrings[i].endsWith(",") || substrings[i].endsWith("”")||substrings[i].endsWith(".") || substrings[i].endsWith('"')){
        //var p = substrings[i].charAt(substrings[i].length-1);
        var q = substrings[i].charAt(substrings[i].length-2); 
        var new_text = removePunctuation(substrings[i]);
        // Logger.log(new_text);
        var wordlength = new_text.length*len;//gets the lengths of each word
        var rect = addRectangle(slidesID, pageID, wordlength);        
        changeRectangleColor(slidesID, rect);
        x = x + wordlength +len +len; 
        if(q == "!" || q == "?" || q == "." || q == ","){
          x = x +len;
        }
      }
      else{
        var wordlength = substrings[i].length*len; //gets the lengths of each word 
        var rect = addRectangle(slidesID, pageID, wordlength);     
        changeRectangleColor(slidesID, rect);
        x = x + wordlength +len;
      }
    Utilities.sleep(1000);}
 // Utilities.sleep(20000);
  }
}


function type4(sentences, page, slides){
  x = 125.5; y = 52;
  var pageID = slides[page-1].objectId; //get the pageID of the slide
  
  for(var w = 0; w < sentences.length; w++){ // loops through the number of sentences
    var text = sentences[w]; //two sentences
    if(w == 1){ //reset starting point
      x = 125.5;
      y = 139;
    }
    else if(w == 2){
      x = 125.5;
      y = 225;
    }
    else if(w == 3){
      x = 125.5;
      y = 312;
    }
    var text = sentences[w]; //whole sentence
    var substrings = text.split(' ');
    
    for(var i = 0; i < substrings.length; i++){
      if(substrings[i].startsWith('¡') || substrings[i].startsWith("¿")|| substrings[i].startsWith("“") || substrings[i].startsWith("/") || substrings[i].startsWith('"')){
        var temp = substrings[i];
        //Logger.log("Start char: " + temp[0]);
        x = x + len;
        if (temp[1] == '¡' || temp[1] == "¿"){
          x = x + len;
        }
        if (substrings[i].endsWith("!") || substrings[i].endsWith("?") ||substrings[i].endsWith(",") ||substrings[i].endsWith("”") ||substrings[i].endsWith(".") || substrings[i].endsWith('"')){
          //var p = substrings[i].charAt(substrings[i].length-1); //p is the last character
          var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character
          //Logger.log("End char: " + p);
          var new_text = removePunctuation(substrings[i]);
          //Logger.log(new_text.length);
          var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths
          
          var rect = addRectangle(slidesID, pageID, wordlength);    
          changeRectangleColor(slidesID, rect);
          
          x = x + wordlength +len +len; //should be + space + size of punctuation;
          if(q == "!" || q == "?" || q == "." || q == ","){
            //Logger.log("2nd to last: " + q);
            x = x + len;
          }   
        }
        else{
          var new_text = removePunctuation(substrings[i]);
          //Logger.log(new_text);
          var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths
          //if type 2:
          //for(var j = 0; j<wordlength.length; j++){
          var rect = addRectangle(slidesID, pageID, wordlength);      
          changeRectangleColor(slidesID, rect);
          x = x + wordlength +len;
        }
      }
      else if (substrings[i].endsWith("?") || substrings[i].endsWith("!") || substrings[i].endsWith(",") || substrings[i].endsWith("”")||substrings[i].endsWith(".") || substrings[i].endsWith('"')){
        var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character

        var new_text = removePunctuation(substrings[i]);
        //Logger.log(new_text);
        var wordlength = new_text.length*len;//gets the lengths of each word //words is an array of the wordlengths
        //Logger.log("Wordlength: " + wordlength);
        //if type 2:
        //for(var j = 0; j<wordlength.length; j++){
        var rect = addRectangle(slidesID, pageID, wordlength);     
        changeRectangleColor(slidesID, rect);
        x = x + wordlength +len +len; //should be + 5 + size of punctuation;
        if(q == "!" || q == "?" || q == "." || q == ","){
          //Logger.log("2nd to last: " + q);
          x = x +len;
        }
      }
      else{
        var wordlength = substrings[i].length*len; //gets the lengths of each word
        var rect = addRectangle(slidesID, pageID, wordlength);      
        changeRectangleColor(slidesID, rect);
        x = x + wordlength +len;
      }
    Utilities.sleep(1000);}
   // Utilities.sleep(20000);
  }
}

/**
  * Creates a Slides API service object and logs the number of slides and
  * elements in a sample presentation:
  * https://docs.google.com/presentation/d/1EAYk18WDjIG-zp_0vLm3CsfQh_i8eXc67Jo2O9C6Vuc/edit
  */

function logProductInfo() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  slidesID = (data[1][4]);
  for (var i = 32; i < data.length; i++) {
    var type = data[i][21];
   // Logger.log("Type:" + data[i][21]);
    if(type == ""){
     // Logger.log("Nothing there");
    }
    else{
      word_arr.push(data[i][6]);
      type_arr.push(data[i][21]);
      page_arr.push(data[i][22]);
     // Logger.log("Added");
    }
  }
Logger.log("Sent: " + word_arr);
 Logger.log("Type: " + type_arr);
 Logger.log("Page: " + page_arr);
}

/**
 * Add a new rectangle to a page.
 * @param {string} presentationId The presentation ID.
 * @param {string} pageId The page ID. 
 * @param {integer} length The rectangle length
 */
function addRectangle(presentationId, pageId, length) {
  // You can specify the ID to use for elements you create,
  // as long as the ID is unique.
  var pageElementId = Utilities.getUuid();

  var requests = [{
    'createShape': {
      'objectId': pageElementId,
      'shapeType': 'ROUND_RECTANGLE',
      'elementProperties': {
        'pageObjectId': pageId,
        'size': {
          'width': {
            'magnitude': length,
            'unit': 'PT'
          },
          'height': {
            'magnitude': 20,
            'unit': 'PT'
          }
        },
        'transform': {
          'scaleX': 1,
          'scaleY': 1,
          'translateX': x,
          'translateY': y,
          'unit': 'PT'
        },            
      }
    }
  }
                 ];
  var response =
      Slides.Presentations.batchUpdate({'requests': requests}, presentationId);
  //Logger.log('Created Textbox with ID: ' + response.replies[0].createShape.objectId);
  return pageElementId;
}

/**
 * Changes the rectangle color and outline.
 * @param {string} presentationId The presentation ID.
 * @param {string} pageElementId The Id of the rectangle
 */

function changeRectangleColor(presentationId, pageElementId) {
  // You can specify the ID to use for elements you create,
  // as long as the ID is unique.

  var requests = [{
    "updateShapeProperties": {
      "objectId": pageElementId,
      "fields": "*",
      "shapeProperties": {
        "shapeBackgroundFill": {
          "solidFill": {
            "color": {
              "themeColor": 'ACCENT4'
              //DARK1 = BLACK
              //DARK2 = DARK GRAY
              //LIGHT1 = WHITE
              //LIGHT2 = DEFAULT GRAY
              //ACCENT1 = ORANGE/YELLOW
              //ACCENT2 = EVEN DARKER GRAY
              //ACCENT3 = TEAL GRAY
              //ACCENT4 = NORMAL GREEN
              //ACCENT5 = TEAL
              //ACCENT6 = BRIGHT YELLOW
            }
          }
        },
        "outline": {
            "dashStyle": "SOLID",
            "outlineFill": {
              "solidFill": {
                "alpha": 1,
                "color": {
                  "themeColor": "ACCENT4"
                }
              }
            },
            "weight": {
              "magnitude": 1,
              "unit": "PT"
            }
      }
    }
          
    }}];
  var response =
      Slides.Presentations.batchUpdate({'requests': requests}, presentationId);
}
                 
/**
 * @see https://remarkablemark.org/blog/2019/09/28/javascript-remove-punctuation/
 */

var regex = /[!¡“#$%&'"”*+,/.:;<=>?¿[\]^_`{|}~]/g;

/**
 * Removes punctuation.
 *
 * @param {string} string
 * @return {string}
 */
function removePunctuation(string) {
  var string1 = string.replace(regex, '');
  var final_string = string1.replace(/\s{2,}/g," ");
  return final_string
}