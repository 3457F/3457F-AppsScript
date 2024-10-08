/**
 * GLOBAL CONSTANTS
 */


/**
 * The standard font size used for page number textboxes
 */
const page_number_font_size = 18;

/**
 * The standard FONT used for page number textboxes
 */
const page_number_font_family = 'Comfortaa';
/**
 * A regex (regular expression) pattern to detect for page number textboxes in a page.
 * 
 * It checks for a procession of digits (a number) followed by any amount of whitespace (usually newlines, because textboxes usually have newlines at the end).
 * In essence, a textbox with JUST a number inside it.
 */
// const pageNumberRegex = new RegExp(
//   "^(\d{1,})" + // checks for JUST a number (ignores other textboxes that MAY have numbers in them, but ALSO text)
//                 // KEEP THE MATCHING GROUP; yes it might reduce efficiency but it plays a purpose!

//   "\s+?$"       // checks for any trailing whitespace (since textboxes have a stupid newline at the end of them)
// );
const pageNumberRegex = /^(\d{1,})\s+?$/;
/**
 * A map of characters, and their respective width values (in pixels) in the Comfortaa font, courtesy of: https://chrishewett.com/blog/calculating-text-width-programmatically/
 */
const letterMapSingle = new Map(
  [[" ",25],["!",33.305],["\"",40.82],["#",50],["$",50],["%",83.305],["&",77.789],["'",18.023],["(",33.305],[")",33.305],["*",50],["+",56.398],[",",25],["-",33.305],[".",25],["/",27.789],["0",50],["1",50],["2",50],["3",50],["4",50],["5",50],["6",50],["7",50],["8",50],["9",50],[":",27.789],[";",27.789],["<",56.398],["=",56.398],[">",56.398],["?",44.391],["@",92.094],["A",72.219],["B",66.703],["C",66.703],["D",72.219],["E",61.086],["F",55.617],["G",72.219],["H",72.219],["I",33.305],["J",38.922],["K",72.219],["L",61.086],["M",88.922],["N",72.219],["O",72.219],["P",55.617],["Q",72.219],["R",66.703],["S",55.617],["T",61.086],["U",72.219],["V",72.219],["W",94.391],["X",72.219],["Y",72.219],["Z",61.086],["[",33.305],["\\",27.789],["]",33.305],["^",46.93],["_",50],["`",33.305],["a",44.391],["b",50],["c",44.391],["d",50],["e",44.391],["f",33.305],["g",50],["h",50],["i",27.789],["j",27.789],["k",50],["l",27.789],["m",77.789],["n",50],["o",50],["p",50],["q",50],["r",33.305],["s",38.922],["t",27.789],["u",50],["v",50],["w",72.219],["x",50],["y",50],["z",44.391],["{",48],["|",20.023],["}",48],["~",54.102],["_median",50]]
);
// .forEach((char_width, char) => {
  
// });
// gets the google slides this script is tied to (in this case, the notebook)
const notebook = SlidesApp.getActivePresentation();
// gets the width / height of slides in the notebook (to calculate alignment for page number textboxes)
const notebook_width = notebook.getPageWidth();
const notebook_height = notebook.getPageHeight();
// get all pages in the notebook
const pages = notebook.getSlides();
// which "presentation page number" (absolute / actual page number) to start "relative" page numbering from
const rel_page_num_start = 3;

// a bunch of arrays used for printing summary info to the screen
let wrong_page_nums = [];
let no_page_nums = [];
let wrong_left_dim = [];
let wrong_top_dim = [];
let wrong_left_pos = [];
let wrong_top_pos = [];
let wrong_font_family = [];
let wrong_font_size = [];

// amount of impostors removed
let impostor_removed_count = 0;


// nerd explanation:
// - each pixel is 1/96th of an inch, and each inch has 72 points in it

// - therefore, we need to divide the pixel value by 96 to get the amount of inches,
// - then multiply by 72 to get the amount of points
// - in other words, (72 * pixels) / 96, or pixels * (72 / 96)
/**
 * Converts a pixel amount to the equivalent amount in "points" (a basic unit of measurement in typography / fonts that Google Slides uses for measurement)
 * 
 * @param {Number} pixels The number, in pixels, to convert
 */
const pixelsToPoints = (pixels) => pixels * (72 / 96);


/**
 * Calculates the width of a string, in "points" (a basic unit of measurement in typography / font that Google Slides uses for measurement)
 * 
 * Courtesy of https://chrishewett.com/blog/calculating-text-width-programmatically/
 * 
 * @param {String} text The text to calculate the width of
 */
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
 * Resizes the `textbox` to fit the text in it (with some padding); also adjusts font size!
 * 
 * @param {SlidesApp.Shape} textbox The textbox to resize
 * @param {Number} curr_page_num_rel The "relative" page number (taking into account skipped pages at the beginning) of the slide the textbox is in (this is for informatics purposes)
 * @param {Boolean} debug_print Whether to do informatics (let the user know WHICH textboxes were resized)
 */
function resizeTextboxToFit(textbox, curr_page_num_rel, informatics = true) {
  // aligns text in textbox to middle, so it isn't wonky when we re-size
  textbox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

  let actual_page_number_font_size = textbox.getText().getTextStyle().getFontSize();

  if (actual_page_number_font_size != page_number_font_size) {
    if (informatics) {
      Logger.log(`Page number textbox has incorrect font size (currently ${actual_page_number_font_size} pts), adjusting...`);

      wrong_font_size.push(curr_page_num_rel);
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

      wrong_left_dim.push(curr_page_num_rel);
    }

    textbox.setWidth(expected_textbox_width);

    Logger.log('Adjusted width!');
  }

  // checks if height is correct; adjusts accordingly
  if (actual_textbox_height != expected_textbox_height) {
    if (informatics) {
      Logger.log(`Height does not match (currently ${actual_textbox_height} pts), adjusting...`);

      wrong_top_dim.push(curr_page_num_rel);
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
 * @param {Number} curr_page_num_rel The "relative" page number (taking into account skipped pages at the beginning) of the slide the textbox is in (this is for informatics purposes)
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

      wrong_left_pos.push(curr_page_num_rel);
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

      wrong_top_pos.push(curr_page_num_rel);
    }

    obj.setTop(obj_expected_top);

    if (informatics) {
      Logger.log('Adjusted top position!');
    }
  }
}

/**
 * Numbers a given page in the notebook (using global variable `pages` which is an array of all the pages in the notebook to number).
 * 
 * Meant to be called in a loop such as `for (let curr_page_idx = 0; curr_page_idx < pages.length; curr_page_idx++)`
 * 
 * With the only config variables being global variables signifying font size (`page_number_font_size`) and font family (`page_number_font_family`), it:
 * - standardizes ALL page numbers with those given sizes and fonts.
 * - resizes the page number textboxes to FIT the text
 * - aligns page number textboxes them at the bottom right of the screen
 * - automatically creates and adds page numbers for slides with missing ones
 * - automatically corrects page numbers for slides with incorrect ones
 * 
 * @param {SlidesApp.Shape} curr_txtbox The textbox to check
 * @param {Number} curr_page_num_rel The "relative" page number of the page (taking into account skipped pages at the beginning)
 * @param {Boolean} page_num_txtbox_exists Whether a textbox has already been found and adjusted accordingly (in which case we are just checking for and deleting duplicates)
 */
function numberPage(curr_txtbox, curr_page_num_rel, page_num_txtbox_exists) {
  // get the text content
  let curr_txtbox_txt_obj = curr_txtbox.getText();
  let curr_txtbox_txt = curr_txtbox_txt_obj.asRenderedString();

  Logger.log(`This textbox's text is: ${curr_txtbox_txt}!`);

  // check text content against regex
  let page_num_match = curr_txtbox_txt.match(pageNumberRegex);

  // if the match succeeded (there is a page number textbox!)
  if (page_num_match) {
    // and it's not a duplicate!
    if (!page_num_txtbox_exists) {
      // get the actual page number
      let page_num_actual = page_num_match[1];

      // lez goo page number is correct
      if (parseInt(page_num_actual, 10) == curr_page_num_rel) {
        Logger.log(`Expected page number: ${curr_page_num_rel}. Current page number is correct; no changes required.`);

      // uh oh! wrong page number!
      } else {
        Logger.log(`Expected page number: ${curr_page_num_rel}. Issue, reconciling page numbers...`);

        wrong_page_nums.push(curr_page_num_rel);

        // sets the textbox to correct number
        curr_txtbox_txt_obj.setText(`${curr_page_num_rel}`);

        Logger.log('Page numbers reconciled!');
      }

      // check if the textbox is correctly aligned and sized
      resizeTextboxToFit(curr_txtbox, curr_page_num_rel);
      alignBottomRight(curr_txtbox, curr_page_num_rel);

      // gets current font family of page number textbox
      let actual_page_number_font_family = curr_txtbox_txt_obj.getTextStyle().getFontFamily();

      if (actual_page_number_font_family != page_number_font_family) {
        Logger.log(`Page number textbox has incorrect font (currently ${actual_page_number_font_family}), adjusting...`);

        wrong_font_family.push(curr_page_num_rel);

        // sets textbox to Comfortaa font (to standardize it with the rest of the notebook)
        curr_shape_txt_obj.getTextStyle().setFontFamily(page_number_font_family);

        Logger.log('Adjusted font family!');
      }
    // BUT IT'S A DUPLICATE
    } else {
      Logger.log("WE HAVE AN IMPOSTOR AMONG US. I CAST THY OUT WITH MY DIVINE POWER!");

      // deletes the impo- i mean duplicate
      curr_txtbox.remove();
      impostor_removed_count++;
      Logger.log(`<textbox> was an impostor. x impostors remain`);
    }

    // either way, a page number textbox was found this run!
    return true;

  // page number not found yet :(
  } else {
    return false;
  }
}

/**
 * @param {SlidesApp.Page} curr_page The page that needs to be checked for correct page numbering
 * @param {Number} curr_page_num_rel The "relative" page number of the page (taking into account skipped pages at the beginning)
 */
function createPageNumber(curr_page, curr_page_num_rel) {
  // uh oh! page number textbox doesn't exist!
  Logger.log('Page number textbox doesn\'t exist, creating...');

  no_page_nums.push(curr_page_num_rel);

  let page_num_textbox = curr_page.insertTextBox(`${curr_page_num_rel}`);

  resizeTextboxToFit(page_num_textbox, curr_page_num_rel, false);
  alignBottomRight(page_num_textbox, curr_page_num_rel, false);

  // sets textbox to Comfortaa font (to standardize it with the rest of the notebook)
  // DOES NOT LOG, because new textboxes are obv gonna be wrong font
  page_num_textbox.getText().getTextStyle().setFontFamily(page_number_font_family);

  Logger.log('Page number textbox created!');
}

function tableOfContents() {

}

/**
 * Fixes various issues with the notebook (such as adding page numbers and keeping the table of contents up to date)
 */
function fixNotebook() {
  // resets the bunch of arrays
  wrong_page_nums = [];
  no_page_nums = [];
  wrong_left_dim = [];
  wrong_top_dim = [];
  wrong_left_pos = [];
  wrong_top_pos = [];

  // loop through every page in the notebook!
  for (let curr_page_idx = 0; curr_page_idx < pages.length; curr_page_idx++) {
  // in case your code doesn't work and you want to run page numbering for only certain slides
  // // for (let curr_page_idx = 0; curr_page_idx < 4; curr_page_idx++) {
    // the list of pages is zero-indexed list, and we're starting from page 3. we add +1 to index to get "actual" page number,
    // then add (1 - 3) (representing the two "skipped" pages) to get OUR "relative" page number, starting from page 3; +1-2 = -1!
    let curr_page_num_rel = curr_page_idx + (1 - (rel_page_num_start - 1));

    Logger.log(`Relative Page #${
      (curr_page_num_rel >= 1) ? curr_page_num_rel : "NOT COUNTED"
    } (abs page idx ${curr_page_idx} || abs page #${curr_page_idx + 1})...`);

    // resets runtime vars
    /**
     * Whether a page number textbox has been found for the current page
     */
    page_num_txtbox_exists = false;

    // fetches the page, using its index!
    let curr_page = pages[curr_page_idx];

    let curr_page_txtboxes = curr_page.getPageElements() // gets all PAGE elements in the page
      .filter((ele) => {                    
        return (
          ele.getPageElementType() === SlidesApp.PageElementType.SHAPE // that are shapes
          && ele.asShape().getShapeType() === SlidesApp.ShapeType.TEXT_BOX       // AND are textboxes
        );
      }).map((ele) => {
        return ele.asShape(); // converts the elements to shapes, NOW THAT WE'RE SURE THEY'RE SHAPES!
      });

    // don't run if this is the first two pages
    if (curr_page_num_rel >= 1) {
      // loop through every textbox in the page!
      for (let txtbox_idx = 0; txtbox_idx < curr_page_txtboxes.length; txtbox_idx++) {
        // fetches the textbox we want to look at 
        let curr_txtbox = curr_page_txtboxes[txtbox_idx];

        // the function returns a boolean signifying whether or not it found a page number textbox
        // if page_num_txtbox_exists is true before the function runs, it will still run
        // but it will just be checking for duplicate page numbers, not ADDING new
        // page numbers
        page_num_txtbox_exists = numberPage(curr_txtbox, curr_page_num_rel, page_num_txtbox_exists);
      }

      Logger.log(page_num_txtbox_exists);

      // didn't find a page number textbox? create one!
      if (!page_num_txtbox_exists) {
        createPageNumber(curr_page, curr_page_num_rel);
      }
    }

    // artifical new line
    Logger.log("");
  }

  // lets user know what issues were fixed
  Logger.log(`NOTE: RELATIVE page numbering starts at page ${rel_page_num_start}!`);
  Logger.log('ISSUES FIXED:');
  Logger.log(`${impostor_removed_count} impostors removed`);
  Logger.log(`Pages with INCORRECT page number: ${!isEmptyFormal(wrong_page_nums) ? wrong_page_nums.join(", ") : "NONE"}`);
  Logger.log(`Pages with NO page number: ${!isEmptyFormal(no_page_nums) ? no_page_nums.join(", ") : "NONE"}`);
  Logger.log(`Page number textboxes with incorrect WIDTH: ${!isEmptyFormal(wrong_left_dim) != [] ? wrong_left_dim.join(", ") : "NONE"}`);
  Logger.log(`Page number textboxes with incorrect HEIGHT: ${!isEmptyFormal(wrong_top_dim) ? wrong_top_dim.join(", ") : "NONE"}`);
  Logger.log(`Page number textboxes in the incorrect HORIZONTAL position: ${!isEmptyFormal(wrong_left_pos) != [] ? wrong_left_pos.join(", ") : "NONE"}`);
  Logger.log(`Page number textboxes in the incorrect VERTICAL position: ${!isEmptyFormal(wrong_top_pos) != [] ? wrong_top_pos.join(", ") : "NONE"}`);
  Logger.log(`Page number textboxes with incorrect font SIZE: ${!isEmptyFormal(wrong_font_size) ? wrong_font_size.join(", ") : "NONE"} `);
  Logger.log(`Page number textboxes with incorrect font FAMILY: ${!isEmptyFormal(wrong_font_family) != [] ? wrong_font_family.join(", ") : "NONE"}`);
}