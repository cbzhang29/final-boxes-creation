var slidesID = ""; //actual slides id
//var slidesID = "1KSRNvuKJRybgqiTNzpipfGVbNttjpcUJPEncX52qkNA";
var word_arr = []; //sentences
var type_arr = []; //types
var page_arr = []; //page num

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

function main(){
  logProductInfo(); // three arrays: word_arr, type_arr, page_arr; should all be the same length
  var presentation = Slides.Presentations.get(slidesID);
  //Logger.log("Sent: " + word_arr);
 // Logger.log("Type: " + type_arr);
 // Logger.log("Page: " + page_arr);
  for(var z = 0; z < word_arr.length; z++){ 
    //Logger.log("z = " + z);
    if (type_arr[z] == 1){
      type1_test(word_arr[z], page_arr[z]);
      Logger.log("Type 1 just ran.");
      Utilities.sleep(7000);
    }
    else if (type_arr[z] == 2){
      var sents_arr = [word_arr[z], word_arr[z+1]];
      type2_test(sents_arr, page_arr[z]);
      z++;
      Logger.log("Type 2 just ran.");
      Utilities.sleep(7000);
    }
    else if (type_arr[z] == 3){
      var sents_arr = [word_arr[z], word_arr[z+1], word_arr[z+2]];
      type3_test(sents_arr, page_arr[z]);
      z = z+2;
      Logger.log("Type 3 just ran.");
      Utilities.sleep(7000);
    }
    else if (type_arr[z] == 4){
      var sents_arr = [word_arr[z], word_arr[z+1], word_arr[z+2], word_arr[z+3]];
      type4_test(sents_arr, page_arr[z]);
      z = z+3;
      Logger.log("Type 4 just ran.");
      Utilities.sleep(7000);
    }
//    else if (type_arr[z] == ''){
//      //don't do anything
//    }
//    else{
//      //Logger.log("Nothing at position " + z);
//    }
  }
}
function test_types(){
  var text = "¿Por qué no / vamos por un café / después? / Test"
  logProductInfo();
  Logger.log(slidesID);
  type1_test(text, 28);
}

function type1_test(sentence, page){
  //x = 360; 
  y = 284; //y should always be 284 //x = 360 is the exact middle
  var presentation = Slides.Presentations.get(slidesID); //get the correct presentation
  var slides = presentation.slides;
  var pageID = slides[page-1].objectId;
  //addRectangle(slidesID, pageID, 10);
  var text = sentence;
  var size = text.length*len;
  var mid = size/2;
  x = 360-mid;
  
  var substrings = text.split(' ');
  Logger.log(substrings);
  //var num_words = substrings.length;
  
  for(var i = 0; i < substrings.length; i++){
    Logger.log("Substring " + i + ": " + substrings[i]);
    if(substrings[i].startsWith('¡') || substrings[i].startsWith("¿")|| substrings[i].startsWith("“")){
      var temp = substrings[i];
      //Logger.log("Start char: " + temp[0]);
      x = x + len;
      if (temp[1] == '¡' || temp[1] == "¿"){
        x = x + len;
      }
      if (substrings[i].endsWith("!") || substrings[i].endsWith("?") ||substrings[i].endsWith(",") ||substrings[i].endsWith("”") ||substrings[i].endsWith(".")){
        //var p = substrings[i].charAt(substrings[i].length-1); //p is the last character
        var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character
        //Logger.log("End char: " + p);
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text.length);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths

        var rect = addRectangle(slidesID, pageID, wordlength);
        changeRectangleColor(slidesID, rect);
        
        x = x + wordlength +len +len; //should be + space + size of punctuation;
        if(q == "!" || q == "?" || q == "." || q == ","){
          Logger.log("2nd to last: " + q);
          x = x + len;
        }
      }
      else{
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths
        //if type 2:
        //for(var j = 0; j<wordlength.length; j++){
        var rect = addRectangle(slidesID, pageID, wordlength);
        changeRectangleColor(slidesID, rect);
        x = x + wordlength +len;
      }
    }
    else if (substrings[i].endsWith("?") || substrings[i].endsWith("!") || substrings[i].endsWith(",") || substrings[i].endsWith("”")||substrings[i].endsWith(".")){
      //var p = substrings[i].charAt(substrings[i].length-1);
      var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character

      //Logger.log("End char: " + p);
      var new_text = removePunctuation(substrings[i]);
      Logger.log(new_text);
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
    else if(substrings[i].startsWith("/") && substrings[i].length == 1){
      var wordlength = substrings[i].length*len; //gets the lengths of each word //words is an array of the wordlengths
      //if type 2:
      //for(var j = 0; j<wordlength.length; j++){
      x = x + wordlength +len;
    }
    else{
      //var new_text = removePunctuation(substrings[i]);
      Logger.log(substrings[i].length);
      var wordlength = substrings[i].length*len; //gets the lengths of each word //words is an array of the wordlengths
      //if type 2:
      //for(var j = 0; j<wordlength.length; j++){
      var rect = addRectangle(slidesID, pageID, wordlength);
        changeRectangleColor(slidesID, rect);
      x = x + wordlength +len;
      //}  
    }
  }
  
  
}

function type2_test(sentences, page){
  var presentation = Slides.Presentations.get(slidesID); //get the correct presentation
  var slides = presentation.slides; //get the slides from the presentation
  //for(var i = 0; i < word_arr.length;i++){ //word_arr.length = # of sentences
  //if type 2:
  x = 125.5; y = 108;
  var pageID = slides[page-1].objectId; //get the pageID of the slide
    //addTextBox(slidesID, pageID, word_arr[i]); //add text to slide
  var text = sentences[0]; //whole sentence
  Logger.log("Text: " + text);
  
  var substrings = text.split(' ');

  for(var i = 0; i < substrings.length; i++){
    Logger.log("Substring " + i + ": " + substrings[i]);
    if(substrings[i].startsWith('¡') || substrings[i].startsWith("¿")|| substrings[i].startsWith("“") || substrings[i].startsWith("/")){
      var temp = substrings[i];
      //Logger.log("Start char: " + temp[0]);
      x = x + len;
      if (temp[1] == '¡' || temp[1] == "¿"){
        x = x + len;
      }
      if (substrings[i].endsWith("!") || substrings[i].endsWith("?") ||substrings[i].endsWith(",") ||substrings[i].endsWith("”") ||substrings[i].endsWith(".")){
        //var p = substrings[i].charAt(substrings[i].length-1); //p is the last character
        var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character
        //Logger.log("End char: " + p);
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text.length);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths

        var rect = addRectangle(slidesID, pageID, wordlength);        
        changeRectangleColor(slidesID, rect);
        
        x = x + wordlength +len +len; //should be + space + size of punctuation;
        if(q == "!" || q == "?" || q == "." || q == ","){
          Logger.log("2nd to last: " + q);
          x = x + len;
        }   
      }
      else{
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths
        //if type 2:
        //for(var j = 0; j<wordlength.length; j++){
        var rect = addRectangle(slidesID, pageID, wordlength);       
        changeRectangleColor(slidesID, rect);
        x = x + wordlength +len;
      }
    }
    else if (substrings[i].endsWith("?") || substrings[i].endsWith("!") || substrings[i].endsWith(",") || substrings[i].endsWith("”")||substrings[i].endsWith(".")){
      //var p = substrings[i].charAt(substrings[i].length-1);
      var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character

      //Logger.log("End char: " + p);
      var new_text = removePunctuation(substrings[i]);
      Logger.log(new_text);
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
      //var new_text = removePunctuation(substrings[i]);
      Logger.log(substrings[i].length);
      var wordlength = substrings[i].length*len; //gets the lengths of each word //words is an array of the wordlengths
      //if type 2:
      //for(var j = 0; j<wordlength.length; j++){
      var rect = addRectangle(slidesID, pageID, wordlength);      
      changeRectangleColor(slidesID, rect);
      x = x + wordlength +len;
      //}  
    }
  }
    x = 125.5;
    y = 251;
    //i++;
    //round two
    //var pageID = slides[4].objectId;     
    var text = sentences[1]; 
    //var text = word_arr[0]; //whole sentence
  Logger.log("Text: " + text);
  
  var substrings = text.split(' ');

  for(var i = 0; i < substrings.length; i++){
    Logger.log("Substring " + i + ": " + substrings[i]);
    if(substrings[i].startsWith('¡') || substrings[i].startsWith("¿")|| substrings[i].startsWith("“") || substrings[i].startsWith("/")){
      var temp = substrings[i];
     // Logger.log("Start char: " + temp[0]);
      x = x +len;
      if (temp[1] == '¡' || temp[1] == "¿"){
        x = x +len;
      }
    
      if (substrings[i].endsWith("!") || substrings[i].endsWith("?") ||substrings[i].endsWith(",") ||substrings[i].endsWith("”") ||substrings[i].endsWith(".")){
        //var p = substrings[i].charAt(substrings[i].length-1); //p is the last character
        var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character
        //Logger.log("End char: " + p);
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text.length);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths

        var rect = addRectangle(slidesID, pageID, wordlength);     
        changeRectangleColor(slidesID, rect);
        
        x = x + wordlength +len +len; //should be + 5 + size of punctuation;
        if(q == "!" || q == "?" || q == "." || q == ","){
          Logger.log("2nd to last: " + q);
          x = x +len;
        }   
      }
      else{
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text.length);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths
        //if type 2:
        //for(var j = 0; j<wordlength.length; j++){
        var rect = addRectangle(slidesID, pageID, wordlength);       
        changeRectangleColor(slidesID, rect);
        x = x + wordlength + len;
      }
    }
    else if (substrings[i].endsWith("?") || substrings[i].endsWith("!") || substrings[i].endsWith(",") || substrings[i].endsWith("”")||substrings[i].endsWith(".")){
      //Logger.log(substrings[i].length);
      //var p = substrings[i].charAt(substrings[i].length-1);
      var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character

      //Logger.log("End char: " + p);
      var new_text = removePunctuation(substrings[i]);
      Logger.log(new_text.length);
      var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths
      //Logger.log("Wordlength: " + wordlength);
      //if type 2:
      //for(var j = 0; j<wordlength.length; j++){
      var rect = addRectangle(slidesID, pageID, wordlength);     
      changeRectangleColor(slidesID, rect);
      x = x + wordlength +len +len; //should be + 5 + size of punctuation;
      if(q == "!" || q == "?" || q == "." || q == ","){
        Logger.log("2nd to last: " + q);
        x = x + len;
      }
    }
    else{
      //var new_text = removePunctuation(substrings[i]);
      Logger.log(substrings[i].length);
      var wordlength = substrings[i].length*len;; //gets the lengths of each word //words is an array of the wordlengths
      //if type 2:
      //for(var j = 0; j<wordlength.length; j++){
      var rect = addRectangle(slidesID, pageID, wordlength);      
      changeRectangleColor(slidesID, rect);
      x = x + wordlength +len;
      //}  
    }
  }
  //}
}

function type3_test(sentences, page){
  var presentation = Slides.Presentations.get(slidesID); //get the correct presentation
  var slides = presentation.slides; //get the slides from the presentation
  //for(var i = 0; i < word_arr.length;i++){ //word_arr.length = # of sentences
  //if type 2:
  x = 125.5; y = 78;
  var pageID = slides[page-1].objectId; //get the pageID of the slide
    //addTextBox(slidesID, pageID, word_arr[i]); //add text to slide
  var text = sentences[0]; //whole sentence
  Logger.log("Text: " + text);
  
  var substrings = text.split(' ');

  for(var i = 0; i < substrings.length; i++){
    if(substrings[i].startsWith('¡') || substrings[i].startsWith("¿")|| substrings[i].startsWith("“") || substrings[i].startsWith("/")){
      var temp = substrings[i];
      //Logger.log("Start char: " + temp[0]);
      x = x + len;
      if (temp[1] == '¡' || temp[1] == "¿"){
        x = x + len;
      }
      if (substrings[i].endsWith("!") || substrings[i].endsWith("?") ||substrings[i].endsWith(",") ||substrings[i].endsWith("”") ||substrings[i].endsWith(".")){
        //var p = substrings[i].charAt(substrings[i].length-1); //p is the last character
        var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character
        //Logger.log("End char: " + p);
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text.length);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths

        var rect = addRectangle(slidesID, pageID, wordlength);    
        changeRectangleColor(slidesID, rect);
        
        x = x + wordlength +len +len; //should be + space + size of punctuation;
        if(q == "!" || q == "?" || q == "." || q == ","){
          Logger.log("2nd to last: " + q);
          x = x + len;
        }   
      }
      else{
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths
        //if type 2:
        //for(var j = 0; j<wordlength.length; j++){
        var rect = addRectangle(slidesID, pageID, wordlength);    
        changeRectangleColor(slidesID, rect);
        x = x + wordlength +len;
      }
    }
    else if (substrings[i].endsWith("?") || substrings[i].endsWith("!") || substrings[i].endsWith(",") || substrings[i].endsWith("”")||substrings[i].endsWith(".")){
      //var p = substrings[i].charAt(substrings[i].length-1);
      var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character

      //Logger.log("End char: " + p);
      var new_text = removePunctuation(substrings[i]);
      Logger.log(new_text);
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
      //var new_text = removePunctuation(substrings[i]);
      Logger.log(substrings[i].length);
      var wordlength = substrings[i].length*len; //gets the lengths of each word //words is an array of the wordlengths
      //if type 2:
      //for(var j = 0; j<wordlength.length; j++){
      var rect = addRectangle(slidesID, pageID, wordlength);     
      changeRectangleColor(slidesID, rect);
      x = x + wordlength +len;
      //}  
    }
  }
    x=125.5
    y = 182;
    //i++;
    //round two    
    var text = sentences[1]; 
    //var text = word_arr[0]; //whole sentence
  Logger.log("Text: " + text);
  
  var substrings = text.split(' ');

  for(var i = 0; i < substrings.length; i++){
    Logger.log("Substring " + i + ": " + substrings[i]);
    if(substrings[i].startsWith('¡') || substrings[i].startsWith("¿")|| substrings[i].startsWith("“") || substrings[i].startsWith("/")){
      var temp = substrings[i];
      //Logger.log("Start char: " + temp[0]);
      x = x + len;
      if (temp[1] == '¡' || temp[1] == "¿"){
        x = x + len;
      }
      if (substrings[i].endsWith("!") || substrings[i].endsWith("?") ||substrings[i].endsWith(",") ||substrings[i].endsWith("”") ||substrings[i].endsWith(".")){
        //var p = substrings[i].charAt(substrings[i].length-1); //p is the last character
        var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character
        //Logger.log("End char: " + p);
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text.length);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths

        var rect = addRectangle(slidesID, pageID, wordlength);       
        changeRectangleColor(slidesID, rect);
        
        x = x + wordlength +len +len; //should be + space + size of punctuation;
        if(q == "!" || q == "?" || q == "." || q == ","){
          Logger.log("2nd to last: " + q);
          x = x + len;
        }   
      }
      else{
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths
        //if type 2:
        //for(var j = 0; j<wordlength.length; j++){
        var rect = addRectangle(slidesID, pageID, wordlength);      
        changeRectangleColor(slidesID, rect);
        x = x + wordlength +len;
      }
    }
    else if (substrings[i].endsWith("?") || substrings[i].endsWith("!") || substrings[i].endsWith(",") || substrings[i].endsWith("”") ||substrings[i].endsWith(".")){
      //var p = substrings[i].charAt(substrings[i].length-1);
      var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character

      //Logger.log("End char: " + p);
      var new_text = removePunctuation(substrings[i]);
      Logger.log(new_text);
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
      //var new_text = removePunctuation(substrings[i]);
      Logger.log(substrings[i].length);
      var wordlength = substrings[i].length*len; //gets the lengths of each word //words is an array of the wordlengths
      //if type 2:
      //for(var j = 0; j<wordlength.length; j++){
      var rect = addRectangle(slidesID, pageID, wordlength);      
      changeRectangleColor(slidesID, rect);
      x = x + wordlength +len;
      //}  
    }
  }
  x=127
    y = 286;
    //i++;
    //round two
    
    var text = sentences[2]; 
    //var text = word_arr[0]; //whole sentence
  Logger.log("Text: " + text);
  
  var substrings = text.split(' ');

  for(var i = 0; i < substrings.length; i++){
    Logger.log("Substring " + i + ": " + substrings[i]);
    if(substrings[i].startsWith('¡') || substrings[i].startsWith("¿")|| substrings[i].startsWith("“") || substrings[i].startsWith("/")){
      var temp = substrings[i];
      //Logger.log("Start char: " + temp[0]);
      x = x + len;
      if (temp[1] == '¡' || temp[1] == "¿"){
        x = x + len;
      }
      if (substrings[i].endsWith("!") || substrings[i].endsWith("?") ||substrings[i].endsWith(",") ||substrings[i].endsWith("”") ||substrings[i].endsWith(".")){
        //var p = substrings[i].charAt(substrings[i].length-1); //p is the last character
        var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character
        //Logger.log("End char: " + p);
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text.length);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths

        var rect = addRectangle(slidesID, pageID, wordlength);       
        changeRectangleColor(slidesID, rect);
        
        x = x + wordlength +len +len; //should be + space + size of punctuation;
        if(q == "!" || q == "?" || q == "." || q == ","){
          Logger.log("2nd to last: " + q);
          x = x + len;
        }   
      }
      else{
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths
        //if type 2:
        //for(var j = 0; j<wordlength.length; j++){
        var rect = addRectangle(slidesID, pageID, wordlength);       
        changeRectangleColor(slidesID, rect);
        x = x + wordlength +len;
      }
    }
    else if (substrings[i].endsWith("?") || substrings[i].endsWith("!") || substrings[i].endsWith(",") || substrings[i].endsWith("”")||substrings[i].endsWith(".")){
      //var p = substrings[i].charAt(substrings[i].length-1);
      var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character

      //Logger.log("End char: " + p);
      var new_text = removePunctuation(substrings[i]);
      Logger.log(new_text);
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
      //var new_text = removePunctuation(substrings[i]);
      Logger.log(substrings[i].length);
      var wordlength = substrings[i].length*len; //gets the lengths of each word //words is an array of the wordlengths
      //if type 2:
      //for(var j = 0; j<wordlength.length; j++){
      var rect = addRectangle(slidesID, pageID, wordlength);      
      changeRectangleColor(slidesID, rect);
      x = x + wordlength +len;
      //}  
    }
  }
}


function type4_test(sentences, page){
  var presentation = Slides.Presentations.get(slidesID); //get the correct presentation
  var slides = presentation.slides; //get the slides from the presentation
  //for(var i = 0; i < word_arr.length;i++){ //word_arr.length = # of sentences
  //if type 2:
  x = 127; y = 52;
  var pageID = slides[page-1].objectId; //get the pageID of the slide
    //addTextBox(slidesID, pageID, word_arr[i]); //add text to slide
  var text = sentences[0]; //whole sentence
  Logger.log("Text: " + text);
  
  var substrings = text.split(' ');

  for(var i = 0; i < substrings.length; i++){
    Logger.log("Substring " + i + ": " + substrings[i]);
    if(substrings[i].startsWith('¡') || substrings[i].startsWith("¿")|| substrings[i].startsWith("“") || substrings[i].startsWith("/")){
      var temp = substrings[i];
      //Logger.log("Start char: " + temp[0]);
      x = x + len;
      if (temp[1] == '¡' || temp[1] == "¿"){
        x = x + len;
      }
      if (substrings[i].endsWith("!") || substrings[i].endsWith("?") ||substrings[i].endsWith(",") ||substrings[i].endsWith("”") ||substrings[i].endsWith(".")){
        //var p = substrings[i].charAt(substrings[i].length-1); //p is the last character
        var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character
        //Logger.log("End char: " + p);
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text.length);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths

        var rect = addRectangle(slidesID, pageID, wordlength);    
        changeRectangleColor(slidesID, rect);
        
        x = x + wordlength +len +len; //should be + space + size of punctuation;
        if(q == "!" || q == "?" || q == "." || q == ","){
          Logger.log("2nd to last: " + q);
          x = x + len;
        }   
      }
      else{
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths
        //if type 2:
        //for(var j = 0; j<wordlength.length; j++){
        var rect = addRectangle(slidesID, pageID, wordlength);      
        changeRectangleColor(slidesID, rect);
        x = x + wordlength +len;
      }
    }
    else if (substrings[i].endsWith("?") || substrings[i].endsWith("!") || substrings[i].endsWith(",") || substrings[i].endsWith("”")||substrings[i].endsWith(".")){
      //var p = substrings[i].charAt(substrings[i].length-1);
      var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character

      //Logger.log("End char: " + p);
      var new_text = removePunctuation(substrings[i]);
      Logger.log(new_text);
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
      //var new_text = removePunctuation(substrings[i]);
      Logger.log(substrings[i].length);
      var wordlength = substrings[i].length*len; //gets the lengths of each word //words is an array of the wordlengths
      //if type 2:
      //for(var j = 0; j<wordlength.length; j++){
      var rect = addRectangle(slidesID, pageID, wordlength);      
      changeRectangleColor(slidesID, rect);
      x = x + wordlength +len;
      //}  
    }
  }
  Utilities.sleep(10000);
    x= 125.5
    y = 139;
    //i++;
    //round two
       
    var text = sentences[1]; 
    //var text = word_arr[0]; //whole sentence
  Logger.log("Text: " + text);
  
  var substrings = text.split(' ');

  for(var i = 0; i < substrings.length; i++){
    Logger.log("Substring " + i + ": " + substrings[i]);
    if(substrings[i].startsWith('¡') || substrings[i].startsWith("¿")|| substrings[i].startsWith("“") || substrings[i].startsWith("/")){
      var temp = substrings[i];
      //Logger.log("Start char: " + temp[0]);
      x = x + len;
      if (temp[1] == '¡' || temp[1] == "¿"){
        x = x + len;
      }
      if (substrings[i].endsWith("!") || substrings[i].endsWith("?") ||substrings[i].endsWith(",") ||substrings[i].endsWith("”") ||substrings[i].endsWith(".")){
        //var p = substrings[i].charAt(substrings[i].length-1); //p is the last character
        var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character
        //Logger.log("End char: " + p);
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text.length);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths

        var rect = addRectangle(slidesID, pageID, wordlength);    
        changeRectangleColor(slidesID, rect);
        
        x = x + wordlength +len +len; //should be + space + size of punctuation;
        if(q == "!" || q == "?" || q == "." || q == ","){
          Logger.log("2nd to last: " + q);
          x = x + len;
        }   
      }
      else{
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths
        //if type 2:
        //for(var j = 0; j<wordlength.length; j++){
        var rect = addRectangle(slidesID, pageID, wordlength);     
        changeRectangleColor(slidesID, rect);
        x = x + wordlength +len;
      }
    }
    else if (substrings[i].endsWith("?") || substrings[i].endsWith("!") || substrings[i].endsWith(",") || substrings[i].endsWith("”") ||substrings[i].endsWith(".")){
      //var p = substrings[i].charAt(substrings[i].length-1);
      var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character

      //Logger.log("End char: " + p);
      var new_text = removePunctuation(substrings[i]);
      Logger.log(new_text);
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
      //var new_text = removePunctuation(substrings[i]);
      Logger.log(substrings[i].length);
      var wordlength = substrings[i].length*len; //gets the lengths of each word //words is an array of the wordlengths
      //if type 2:
      //for(var j = 0; j<wordlength.length; j++){
      var rect = addRectangle(slidesID, pageID, wordlength);     
      changeRectangleColor(slidesID, rect);
      x = x + wordlength +len;
      //}  
    }
  }
  Utilities.sleep(10000);
  x= 125.5
  y = 225;
    //i++;
    //round two
      
    var text = sentences[2]; 
    //var text = word_arr[0]; //whole sentence
  Logger.log("Text: " + text);
  
  var substrings = text.split(' ');

  for(var i = 0; i < substrings.length; i++){
    Logger.log("Substring " + i + ": " + substrings[i]);
    if(substrings[i].startsWith('¡') || substrings[i].startsWith("¿")|| substrings[i].startsWith("“") || substrings[i].startsWith("/")){
      var temp = substrings[i];
      //Logger.log("Start char: " + temp[0]);
      x = x + len;
      if (temp[1] == '¡' || temp[1] == "¿"){
        x = x + len;
      }
      if (substrings[i].endsWith("!") || substrings[i].endsWith("?") ||substrings[i].endsWith(",") ||substrings[i].endsWith("”") ||substrings[i].endsWith(".")){
        //var p = substrings[i].charAt(substrings[i].length-1); //p is the last character
        var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character
        //Logger.log("End char: " + p);
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text.length);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths

        var rect = addRectangle(slidesID, pageID, wordlength);     
        changeRectangleColor(slidesID, rect);
        
        x = x + wordlength +len +len; //should be + space + size of punctuation;
        if(q == "!" || q == "?" || q == "." || q == ","){
          Logger.log("2nd to last: " + q);
          x = x + len;
        }   
      }
      else{
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths
        //if type 2:
        //for(var j = 0; j<wordlength.length; j++){
        var rect = addRectangle(slidesID, pageID, wordlength);   
        changeRectangleColor(slidesID, rect);
        x = x + wordlength +len;
      }
    }
    else if (substrings[i].endsWith("?") || substrings[i].endsWith("!") || substrings[i].endsWith(",") || substrings[i].endsWith("”") ||substrings[i].endsWith(".")){
      //var p = substrings[i].charAt(substrings[i].length-1);
      var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character

      //Logger.log("End char: " + p);
      var new_text = removePunctuation(substrings[i]);
      Logger.log(new_text);
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
      //var new_text = removePunctuation(substrings[i]);
      Logger.log(substrings[i].length);
      var wordlength = substrings[i].length*len; //gets the lengths of each word //words is an array of the wordlengths
      //if type 2:
      //for(var j = 0; j<wordlength.length; j++){
      var rect = addRectangle(slidesID, pageID, wordlength);    
      changeRectangleColor(slidesID, rect);
      x = x + wordlength +len;
      //}  
    }
  }
  Utilities.sleep(10000);
  x = 125.5
    y = 312;
    //i++;
    //round two
   // var pageID = slides[10].objectId;     
    var text = sentences[3]; 
    //var text = word_arr[0]; //whole sentence
  Logger.log("Text: " + text);
  
  var substrings = text.split(' ');

  for(var i = 0; i < substrings.length; i++){
    Logger.log("Substring " + i + ": " + substrings[i]);
    if(substrings[i].startsWith('¡') || substrings[i].startsWith("¿")|| substrings[i].startsWith("“") || substrings[i].startsWith("/")){
      var temp = substrings[i];
      //Logger.log("Start char: " + temp[0]);
      x = x + len;
      if (temp[1] == '¡' || temp[1] == "¿"){
        x = x + len;
      }
      if (substrings[i].endsWith("!") || substrings[i].endsWith("?") ||substrings[i].endsWith(",") ||substrings[i].endsWith("”") ||substrings[i].endsWith(".")){
        //var p = substrings[i].charAt(substrings[i].length-1); //p is the last character
        var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character
        //Logger.log("End char: " + p);
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text.length);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths

        var rect = addRectangle(slidesID, pageID, wordlength);   
        changeRectangleColor(slidesID, rect);
        
        x = x + wordlength +len +len; //should be + space + size of punctuation;
        if(q == "!" || q == "?" || q == "." || q == ","){
          Logger.log("2nd to last: " + q);
          x = x + len;
        }   
      }
      else{
        var new_text = removePunctuation(substrings[i]);
        Logger.log(new_text);
        var wordlength = new_text.length*len; //gets the lengths of each word //words is an array of the wordlengths
        //if type 2:
        //for(var j = 0; j<wordlength.length; j++){
        var rect = addRectangle(slidesID, pageID, wordlength);     
        changeRectangleColor(slidesID, rect);
        x = x + wordlength +len;
      }
    }
    else if (substrings[i].endsWith("?") || substrings[i].endsWith("!") || substrings[i].endsWith(",") || substrings[i].endsWith("”")||substrings[i].endsWith(".")){
      //var p = substrings[i].charAt(substrings[i].length-1);
      var q = substrings[i].charAt(substrings[i].length-2); // q is the second to last character

      //Logger.log("End char: " + p);
      var new_text = removePunctuation(substrings[i]);
      Logger.log(new_text);
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
      //var new_text = removePunctuation(substrings[i]);
      Logger.log(substrings[i].length);
      var wordlength = substrings[i].length*len; //gets the lengths of each word //words is an array of the wordlengths
      //if type 2:
      //for(var j = 0; j<wordlength.length; j++){
      var rect = addRectangle(slidesID, pageID, wordlength);     
      changeRectangleColor(slidesID, rect);
      x = x + wordlength +len;
      //}  
    }
  }
  Utilities.sleep(10000);
}

/**
  * Creates a Slides API service object and logs the number of slides and
  * elements in a sample presentation:
  * https://docs.google.com/presentation/d/1EAYk18WDjIG-zp_0vLm3CsfQh_i8eXc67Jo2O9C6Vuc/edit
  */

function logProductInfo() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  slidesID = data[1][4];
  for (var i = 32; i < data.length; i++) {
   // Logger.log('Product name: ' + data[i][0]); //row i col 1 (A)
    //Logger.log(data[i][0]); //row i col 2 (B) 
    word_arr.push(data[i][6]);
    type_arr.push(data[i][21]);
    page_arr.push(data[i][22]);
  }
}

/**
 * Add a new rectangle to a page.
 * @param {string} presentationId The presentation ID.
 * @param {string} pageId The page ID.
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
              //ACCENT4 = NORMAL GREEN (same as 1)
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

var regex = /[!¡“#$%&'”()*+,/.:;<=>?¿@[\]^_`{|}~]/g;

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
function testPunct(){
  var new_string = removePunctuation('¿Cómo? Yes, let’s / do it, “she own’s it” /  ¡No way! Saliga-smith.');
  Logger.log(new_string);
}