import java.io.File;
import java.util.Date;
import jxl.*; 

//make a font variable
PFont font;
String thename = "United Kingdom";
int begin = 1;

int txtHeight = 25;
Workbook[] workbook = new Workbook[24];
Sheet[] sheet = new Sheet[24];
Cell[] stats = new Cell[24];
String[] stat = new String[24];
Cell[] cellSheet = new Cell[24];
int[] cellSheetRow = new int[24];
Cell[] datas = new Cell[24];
Cell[] ranks = new Cell[24];
String[] rank = new String[24];
String[] data = new String[24];

void setup() {
  size(600,800);
  background(0);

  //load the font to use
  font = loadFont("Courier-12.vlw"); 

  try {
    for(int i = 1; i <= 23; i++) {
      workbook[i] = Workbook.getWorkbook(new File(sketchPath(i + ".xls")));
      //println("Workbook " + i);
    }
  } 
  catch(Exception e) {
    println("Error! " + e); 
  }


}


void draw() {
  if(begin == 1) {
    background(0);

    fill(0,255,0);
    textFont(font, 16);
    String j = ("Country: " + thename);
    text(j, 25, txtHeight, 650, 775);
    txtHeight += 30;

    textFont(font);
    String cmd = "say " + thename;
    Process p;
    try{
      p = Runtime.getRuntime().exec(cmd);
    }
    catch(IOException e){ 
      e.printStackTrace();
    }


    for(int i = 1; i<= 23; i++) {
      try {
        sheet[i] = workbook[i].getSheet(0);  
        stats[i] = sheet[i].getCell(2,1);                // What is the stat we're looking at?
        stat[i] = stats[i].getContents();
      } 
      catch(Exception n){
        print("Problem with: " + n); 
      }
    }

    delay(20);

    //First COLUMN = 0 (A in a spreadsheet)
    //First ROW    = 0 (1 in a spreadsheet)
    //SO the first COLUMN and ROW = 0,0

      // try {

    for(int i = 1; i<= 23; i++) {

      try {
        cellSheet[i] = sheet[i].findLabelCell(thename);
        cellSheetRow[i] = cellSheet[i].getRow();

        delay(20);

        datas[i] = sheet[i].getCell(2, cellSheetRow[i]);                // Get the ranking number
        ranks[i] = sheet[i].getCell(0, cellSheetRow[i]);                // Get the stat data

        rank[i] = ranks[i].getContents();
        data[i] = datas[i].getContents(); 

        //println(stat[i] + ": ");
        //println(thename + " is ranked " + rank[i] + " with " + data[i]); 
        //println(" ");

        // write the font
        String m = (stat[i] + ": ");
        text(m, 25, txtHeight, 650, 775);
        txtHeight += 12;

        String s = ("Ranked " + rank[i] + " with " + data[i]);
        text(s, 25, txtHeight, 650, 775);
        txtHeight += 20;
      } 

      catch(Exception l) {
        String m = (stat[i] + ": ");
        text(m, 25, txtHeight, 650, 775);
        txtHeight += 12;
        String s = ("not available");
        text(s, 25, txtHeight, 650, 775);
        txtHeight += 20;

        println(stat[i] + ": ");
        println("Problem: " + l + " in workbook " + i); 
        println(" "); 
      }
    }
    //}


    for(int i = 1; i<= 23; i++) {
      workbook[i].close();
    }
    delay(20);
    begin = 0;
  } 
  else {
    voice();
    noLoop();
  }

}





void voice() {
  String sayThis = "";
  for (int i=1; i<=23; i++) {
    //sayThis = sayThis + stat[i] + "Ranked " + rank[i] + " with " + data[i] + ". ";
    sayThis = "Country Statistics go here. ";
  }

  sayThis = sayThis + " Process Complete.";

  String cmd1 = "say " + sayThis;
  
  // saveStrings("lines.txt", lines);   http://processing.org/learning/topics/savefile1.html
  // http://processing.org/learning/topics/savefile2.html
  //       workbook[i] = Workbook.getWorkbook(new File(sketchPath(i + ".xls")));
  
  String cmd2 = "lpr " + sayThis;
  
  Process o;
  Process r;
  try{
    o = Runtime.getRuntime().exec(cmd1);
    r = Runtime.getRuntime().exec(cmd2);
  }
  catch(IOException v){ 
    v.printStackTrace();
  }

}
