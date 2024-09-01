/**
 * GLOBAL CONSTANTS
 */
// font size for page number textboxes
const page_number_font_size = 18;
// FONT for page number textboxes
const page_number_font_family = 'Comfortaa';
/**
 * regex pattern to detect for page number textboxes in a page
 * basically checks for any number of digits, followed by any amount of whitespace (bc textboxes have a stupid newline at the end of them)
 * we're detecting all "page number" elements, so the ONLY TEXT in these textboxes should be a number!
 * KEEP THE MATCH IN THIS REGEX; IT'S NOT USELESS AND ACTUALLY HAS A PURPOSE
 */
const pageNumberRegex = /^(\d{1,})\s+?$/;
// width and kerning values for Comfortaa font, courtesy of: https://chrishewett.com/blog/calculating-text-width-programmatically/
const letterMapSingle = new Map([[" ",25],["!",33.305],["\"",40.82],["#",50],["$",50],["%",83.305],["&",77.789],["'",18.023],["(",33.305],[")",33.305],["*",50],["+",56.398],[",",25],["-",33.305],[".",25],["/",27.789],["0",50],["1",50],["2",50],["3",50],["4",50],["5",50],["6",50],["7",50],["8",50],["9",50],[":",27.789],[";",27.789],["<",56.398],["=",56.398],[">",56.398],["?",44.391],["@",92.094],["A",72.219],["B",66.703],["C",66.703],["D",72.219],["E",61.086],["F",55.617],["G",72.219],["H",72.219],["I",33.305],["J",38.922],["K",72.219],["L",61.086],["M",88.922],["N",72.219],["O",72.219],["P",55.617],["Q",72.219],["R",66.703],["S",55.617],["T",61.086],["U",72.219],["V",72.219],["W",94.391],["X",72.219],["Y",72.219],["Z",61.086],["[",33.305],["\\",27.789],["]",33.305],["^",46.93],["_",50],["`",33.305],["a",44.391],["b",50],["c",44.391],["d",50],["e",44.391],["f",33.305],["g",50],["h",50],["i",27.789],["j",27.789],["k",50],["l",27.789],["m",77.789],["n",50],["o",50],["p",50],["q",50],["r",33.305],["s",38.922],["t",27.789],["u",50],["v",50],["w",72.219],["x",50],["y",50],["z",44.391],["{",48],["|",20.023],["}",48],["~",54.102],["_median",50]]);

let notebook_width = 0;
let notebook_height = 0;

// a bunch of arrays used for printing summary info to the screen
let wrong_page_nums = [];
let no_page_nums = [];
let wrong_left_dim = [];
let wrong_top_dim = [];
let wrong_left_pos = [];
let wrong_top_pos = [];
let wrong_font_family = [];
let wrong_font_size = [];

let testVar = 0;

// 96 pixels per inch ; 72 points per inch
// pixels -> inches = /96
// inches -> points = *72

// pixels -> inches -> points = (pixels / 96) * 72 = 72(pixels)/96 
// so we multiply 

function pixelsToPoints(pixels) {
  return pixels * (72 / 96);
}

// courtesy of https://chrishewett.com/blog/calculating-text-width-programmatically/
function calcTextWidthInPoints(text) {
  let stringWidth = 0;

  const stringSplit = [...text];

  for (const letter of stringSplit) {
    stringWidth += letterMapSingle.get(letter) || letterMapSingle.get('_median');
  }

  // convert to 18px font
  let font_ratio = page_number_font_size / 100;
  stringWidth *= font_ratio;

  // convert string width from pixels to points!
  stringWidth = pixelsToPoints(stringWidth);

  return stringWidth;
}

/**
 * Checks whether the given array is EMPTY (using a formal, "fool-proof" approach as explained in
 * https://stackoverflow.com/a/24403771)
 * 
 * @param {Array} array The array to check
 * @returns {Boolean} Whether the array is empty (true) or not (false)
 * 
 */
function isEmptyFormal(array) {
  if (!Array.isArray(array) || !array.length) {
    return true;
  } else {
    return false;
  }
}

/**
 * Resizes the `textbox` to fit the text in it (with some padding)
 * 
 * @param {SlidesApp.Shape} textbox The textbox to resize
 * @param {Number} curr_page_num The page number of the slide the textbox is in (this is for informatics purposes)
 * @param {Boolean} debug_print Whether to do informatics (let the user know WHICH textboxes were resized)
 */
function resizeTextboxToFit(textbox, curr_page_num, informatics = true) {
  // aligns text in textbox to middle, so it isn't wonky when we re-size
  textbox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

  let actual_page_number_font_size = textbox.getText().getTextStyle().getFontSize();

  if (actual_page_number_font_size != page_number_font_size) {
    if (informatics) {
      Logger.log(`Page number textbox has incorrect font size (currently ${actual_page_number_font_size} pts), adjusting...`);

      wrong_font_size.push(curr_page_num);
    }

    // standardizes font size
    textbox.getText().getTextStyle().setFontSize(page_number_font_size);  

    if (informatics) {
      Logger.log('Adjusted font size!');
    }
  }

  // approximates width of characters as 0.6x font size; sets width to length of text * 
  textbox_text = textbox.getText().asRenderedString();
  // rounding because google slides is crazy weird and decides to use 2394023490823 decimal places for position numbers
  // using "+" to convert .toFixed() result back to number
  expected_textbox_width = +(calcTextWidthInPoints(textbox_text) * 2).toFixed(1);
  // sets height to font size (bc font size is proportional to character height), with 1.2x (~3.6pt) padding
  // rounding because google slides is crazy weird and decides to use 2394023490823 decimal places for position numbers
  // using "+" to convert .toFixed() result back to number
  expected_textbox_height = +(page_number_font_size * 2).toFixed(1);

  Logger.log(`Checking textbox size; expecting Width: ${expected_textbox_width} pts; Height: ${expected_textbox_height} pts`);

  // rounding because google slides is crazy weird and decides to use 2394023490823 decimal places for position numbers
  // using "+" to convert .toFixed() result back to number
  actual_textbox_width = +textbox.getWidth().toFixed(1);
  actual_textbox_height = +textbox.getHeight().toFixed(1);

  // checks if width is correct; adjusts accordingly
  if (actual_textbox_width != expected_textbox_width) {
    if (informatics) {
      Logger.log(`Width does not match (currently ${actual_textbox_width} pts), adjusting...`);

      wrong_left_dim.push(curr_page_num);
    }

    textbox.setWidth(expected_textbox_width);

    Logger.log('Adjusted width!');
  }

  // checks if height is correct; adjusts accordingly
  if (actual_textbox_height != expected_textbox_height) {
    if (informatics) {
      Logger.log(`Height does not match (currently ${actual_textbox_height} pts), adjusting...`);

      wrong_top_dim.push(curr_page_num);
    }

    textbox.setHeight(expected_textbox_height);

    if (informatics) {
      Logger.log('Adjusted height!');
    }
  }
}

/**
 * Aligns the `obj` shape at the bottom right of the presentation
 * 
 * @param {SlidesApp.Shape} obj The object to align
 * @param {Number} curr_page_num The page number of the slide the textbox is in (this is for informatics purposes)
 * @param {Boolean} debug_print Whether to do informatics (let the user know WHICH textboxes were resized)
 */
function alignBottomRight(obj, curr_page_num, informatics = true) {
  // gets object dimensions
  obj_width = obj.getWidth();
  obj_height = obj.getHeight();

  // calculates expected position based on object dimensions
  // rounding because google slides is crazy weird and decides to use 2394023490823 decimal places for position numbers
  // using "+" to convert .toFixed() result back to number
  obj_expected_left = (notebook_width - obj_width).toFixed(1);
  obj_expected_top = (notebook_height - obj_height).toFixed(1);

  Logger.log(`Checking if textbox is correctly aligned; expecting position Left: ${obj_expected_left} pts; Top: ${obj_expected_top} pts`);

  // gets actual object position
  // rounding because google slides is crazy weird and decides to use 2394023490823 decimal places for position numbers
  // using "+" to convert .toFixed() result back to number
  obj_actual_left = obj.getLeft().toFixed(1);
  obj_actual_top = obj.getTop().toFixed(1);

  // check if left position is correct; if not, adjusts it accordingly
  if (obj_actual_left != obj_expected_left) {
    if (informatics) {
      Logger.log(`Left position does not match (currently ${obj_actual_left} pts), adjusting...`);

      wrong_left_pos.push(curr_page_num);
    }

    obj.setLeft(obj_expected_left);

    if (informatics) {
      Logger.log('Adjusted left position!');
    }
  }

  // check if top position is correct; if not, adjusts it accordingly
  if (obj_actual_top != obj_expected_top) {
    if (informatics) {
      Logger.log(`Top position does not match (currently ${obj_actual_top} pts), adjusting...`);

      wrong_top_pos.push(curr_page_num);
    }

    obj.setTop(obj_expected_top);

    if (informatics) {
      Logger.log('Adjusted top position!');
    }
  }
}

/**
 * Numbers all the pages in a Google Slides presentation
 * 
 * With the only config variables being global variables signifying font size (`page_number_font_size`) and font family (`page_number_font_family`), it:
 * - standardizes ALL page numbers with those given sizes and fonts.
 * - resizes the page number textboxes to FIT the text
 * - aligns page number textboxes them at the bottom right of the screen
 * - automatically creates and adds page numbers for slides with missing ones
 * - automatically corrects page numbers for slides with incorrect ones
 */
function numberPages() {
  // resets the bunch of arrays
  wrong_page_nums = [];
  no_page_nums = [];
  wrong_left_dim = [];
  wrong_top_dim = [];
  wrong_left_pos = [];
  wrong_top_pos = [];
  

  // gets the google slides this script is tied to (in this case, the notebook)
  const notebook = SlidesApp.getActivePresentation();

  // initializes the width / height of slides in the notebook (to calculate alignment for page number textboxes)
  notebook_width = notebook.getPageWidth();
  notebook_height = notebook.getPageHeight();

  // get all pages in the notebook
  const pages = notebook.getSlides();

  // loop through every page in the notebook!
  for (let page_idx = 0; page_idx < pages.length; page_idx++) {
    // NOT NUMBERING first two pages!
    if (page_idx < 2) {
      continue;
    }

    Logger.log(`Page ${page_idx + 1}...`);

    // boolean variable indicating if a page number textbox has been created for this page
    let page_num_shape_exists = false;

    // the list of pages is zero-indexed list, while page numbers typically start from 1 (at least on Earth)
    // , so add 1 to reconcile the two indexing systems
    // starting from page 3; we add +1 to index to get "actual" page number, then subtract by 2 to get OUR "relative" page number; +1-2 = -1!
    let curr_page_num = page_idx - 1;

    // fetches the page
    let curr_page = pages[page_idx];

    // get all elements in the given page, filtering ONLY for shapes (yes, this does include other, irrelevant elements like rectangles, ellipses, arrows, etc. but the important part is that textboxes are also included in the "shape" classification)
    let curr_page_shapes = curr_page.getPageElements().filter((ele) => {
      return ele.getPageElementType() == SlidesApp.PageElementType.SHAPE;
    });

    // loop through every shape in the page!
    for (let shape_idx = 0; shape_idx < curr_page_shapes.length; shape_idx++) {
      // fetches the shape
      let curr_shape = curr_page_shapes[shape_idx].asShape();

      // get the text content
      let curr_shape_txt_obj = curr_shape.getText();
      let curr_shape_txt = curr_shape_txt_obj.asString();

      // check text content against regex
      let page_num_match = curr_shape_txt.match(pageNumberRegex);

      // if the match succeeded
      if (page_num_match) {
        page_num_shape_exists = true;

        // get the actual page number
        let page_num = page_num_match[1];
        Logger.log(`Current page number: ${curr_page_num}`);

        // lez goo page number is correct
        if (parseInt(page_num, 10) == curr_page_num) {
          Logger.log(`Expected page number: ${curr_page_num}. Current page number is correct; no changes required.`);

        // uh oh! wrong page number!
        } else {
          Logger.log(`Expected page number: ${curr_page_num}. Issue, reconciling page numbers...`);

          wrong_page_nums.push(curr_page_num);

          // sets the textbox to correct number
          curr_shape_txt_obj.setText(`${curr_page_num}`);

          Logger.log('Page numbers reconciled!');
        }

        // check if the textbox is correctly aligned and sized
        resizeTextboxToFit(curr_shape, curr_page_num);
        alignBottomRight(curr_shape, curr_page_num);

        // gets current font family of page number textbox
        let actual_page_number_font_family = curr_shape_txt_obj.getTextStyle().getFontFamily();

        if (actual_page_number_font_family != page_number_font_family) {
          Logger.log(`Page number textbox has incorrect font (currently ${actual_page_number_font_family}), adjusting...`);

          wrong_font_family.push(curr_page_num);

          // sets textbox to Comfortaa font (to standardize it with the rest of the notebook)
          curr_shape_txt_obj.getTextStyle().setFontFamily(page_number_font_family);

          Logger.log('Adjusted font family!');
        }

        // don't need to examine any other shapes in the slide...
        break;
      }
    }

    // uh oh! page number textbox doesn't exist!
    if (!page_num_shape_exists) {
      Logger.log('Page number textbox doesn\'t exist, creating...');

      no_page_nums.push(curr_page_num);

      let page_num_textbox = curr_page.insertTextBox(`${curr_page_num}`);

      resizeTextboxToFit(page_num_textbox, curr_page_num, false);
      alignBottomRight(page_num_textbox, curr_page_num, false);

      // sets textbox to Comfortaa font (to standardize it with the rest of the notebook)
      // DOES NOT LOG, because new textboxes are obv gonna be wrong font
      page_num_textbox.getText().getTextStyle().setFontFamily(page_number_font_family);
    
      Logger.log('Page number textbox created!');
    }

    // artifical new line
    Logger.log("");
  }

  // 1 - wrong font ; 2 - wrong size ; 3 - wrong font AND size ; 4 - wrong horiz pos ; 5 - wrong vert pos
  // 6 - wrong horiz AND vert pos ; 7 - wrong width ; 8 - wrong height ; 9 - wrong width AND height
  // 10 - wrong page number ; 11 - NO page numbers

  Logger.log("ISSUES FIXED:");
  Logger.log(isEmptyFormal(wrong_page_nums));
  Logger.log(`Pages with INCORRECT page number: ${!isEmptyFormal(wrong_page_nums) ? wrong_page_nums.join(", ") : "NONE"}`);
  Logger.log(`Pages with NO page number: ${!isEmptyFormal(no_page_nums) ? no_page_nums.join(", ") : "NONE"}`);
  Logger.log(`Page number textboxes with incorrect WIDTH: ${!isEmptyFormal(wrong_left_dim) != [] ? wrong_left_dim.join(", ") : "NONE"}`);
  Logger.log(`Page number textboxes with incorrect HEIGHT: ${!isEmptyFormal(wrong_top_dim) ? wrong_top_dim.join(", ") : "NONE"}`);
  Logger.log(`Page number textboxes in the incorrect HORIZONTAL position: ${!isEmptyFormal(wrong_left_pos) != [] ? wrong_left_pos.join(", ") : "NONE"}`);
  Logger.log(`Page number textboxes in the incorrect VERTICAL position: ${!isEmptyFormal(wrong_top_pos) != [] ? wrong_top_pos.join(", ") : "NONE"}`);
  Logger.log(`Page number textboxes with incorrect font SIZE: ${!isEmptyFormal(wrong_font_size) ? wrong_font_size.join(", ") : "NONE"} `);
  Logger.log(`Page number textboxes with incorrect font FAMILY: ${!isEmptyFormal(wrong_font_family) != [] ? wrong_font_family.join(", ") : "NONE"}`);
}