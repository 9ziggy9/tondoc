const LETTER_CONTENT = "Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut laoreet dolore magna aliquam erat volutpat. Ut wisi enim ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat.\n\nDuis autem vel eum iriure dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan.\n\nNam liber tempor cum soluta nobis eleifend option congue nihil imperdiet doming id quod mazim placerat facer possim assum. Typi non habent claritatem insitam; est usus legentis in iis qui facit eorum claritatem. Investigationes demonstraverunt lectores legere me lius quod ii legunt saepius.\n\nSincerely,\n";

const DEBUG_DATA = {
  "contact": {
    "first": "David",
    "last": "Rogers",
    "street": "199 Testing Lane",
    "geo": "Chico CA, 95973",
    "phone": "(444) 359-3502",
    "email": "lollercopter@gmail.ruse",
  },
  "clients": [
    {
      "first": "Brock", "last": "Purdy",
      "title": "Quarterback", "comp": "49ers",
      "street": "4949 Niners Way",
      "geo": "San Francisco CA, 99359",
    },
    {
      "first": "Tom", "last": "Brady",
      "title": "Retired", "comp": "Patriots",
      "street": "99 Patriots Street",
      "geo": "Boston MA, 42359",
    },
    {
      "first": "Frankie", "last": "Gore",
      "comp": "49ers",
      "street": "9924 Forty-Niners Way",
    },
  ]
};

// ENVIRONMENT
const LOG_SHEET_ID = SpreadsheetApp
  .getActiveSpreadsheet()
  .getId();
const LOG_SHEET_NAME      = "LogSheet";
const LOG_LABEL_END       = 5;
const LOG_NUM_DATA_POINTS = 7;
const LOG_PHONE_CELL      = "A1";
const LOG_STREET_CELL     = "A2";
const LOG_GEO_CELL        = "A3";
const LOG_EMAIL_CELL      = "A4";
const LOG_COMP_CELL       = "C1";
const LOG_NAME_CELL       = "G2";

// STYLING METHODS
const FONT_FAMILY     = "Proxima Nova";
const FONT_SIZE_SMALL = 10;
const FONT_SIZE_MED   = 11;
const FONT_SIZE_LARGE = 16;

const STYLE_DEFAULT = {
  [DocumentApp.Attribute.BOLD]: false,
  [DocumentApp.Attribute.FONT_FAMILY]: FONT_FAMILY,
  [DocumentApp.Attribute.FONT_SIZE]: FONT_SIZE_SMALL,
};
const STYLE_BOLD = {
  [DocumentApp.Attribute.BOLD]: true,
  [DocumentApp.Attribute.FONT_FAMILY]: FONT_FAMILY,
  [DocumentApp.Attribute.FONT_SIZE]: FONT_SIZE_MED,
};
const STYLE_LARGE = {
  [DocumentApp.Attribute.BOLD]: true,
  [DocumentApp.Attribute.FONT_FAMILY]: FONT_FAMILY,
  [DocumentApp.Attribute.FONT_SIZE]: FONT_SIZE_LARGE,
};