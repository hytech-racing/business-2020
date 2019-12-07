//Other imports
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;
import java.util.Scanner;

//imports for Microsoft Excel
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Makes an Excel Spreadsheet that finds differences between
 * the FY20 sheet and the HyTech Racing Sheet
 * @author Rohit (Rohsomeness B)
 * @version 1.0.0
 */
public class HytechFinanceCleanUp {

    private static ArrayList<Double> HyTechList = new ArrayList<Double>();
    private static ArrayList<Double> FY20List = new ArrayList<Double>();
    private static ArrayList<String> FY20DescList = new ArrayList<String>();
    private static String fy20File = "";
    private static String hyTechFile = "";
    private static String confFile = "";
    private static int counter3 = 1;
    private static int counter4 = 1;
    private static int counter5 = 0;
    private static boolean yeet = false;
    private static Scanner scan = new Scanner(System.in);

    /**
     * Main method. Prompts user with directions
     * @param args what is inputted into the compiler
     */
    public static void main(String[] args) {
        System.out.println("Please confirm that in FY20 file the second sheet column E has lots of "
            + "numbers and that in HyTech File first sheet column C has lots of numbers.");
        System.out.println("Would you like to input file locations?(y/n)");
        if (scan.next().equalsIgnoreCase("y")) {
            System.out.println("Make sure file names do not have spaces in them or an error might be thrown.");
            System.out.println("Enter file location for \"FY20 Finances\" File. Make sure to type in using \"\\\"");
            fy20File = scan.next();
            System.out.println("Enter file location of \"HyTech Racing\" file. Make sure to type in using \"\\\".");
            hyTechFile = scan.next();
            System.out.println("What folder should the confirmation spread sheet be put in? Add a \\ to the end of this file location.");
            confFile = scan.next();

            findMatches(fy20File, hyTechFile, confFile);
        } else {
            System.out.println("Make sure the files are titled \"FY20Finances.xlsx\" and \"HyTechRacing.xlsx\" and are" +
                    " in the folder in which this program is running in. The confirmation "
                + "program will also be put in this folder.");

            findMatches("FY20Finances.xlsx", "HyTechRacing.xlsx", "");
        }

    }

    /**
     * Actual method that does the logic
     * @param fy20File FY20 File (from HyTech Team)
     * @param hyTechFile HyTech File (from GT)
     * @param confFile Confirmation File created to find matches
     */
    public static void findMatches(String fy20File, String hyTechFile, String confFile) {
        try {
            //Makes an arraylist with the info needed from Hytech File
            FileInputStream HyTechFile = new FileInputStream(new File(hyTechFile));
            XSSFWorkbook HyTechWbk = new XSSFWorkbook(HyTechFile);
            XSSFSheet HyTechSheet = HyTechWbk.getSheetAt(0);
            Iterator<Row> rowIterator = HyTechSheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    if ((cell.getColumnIndex() == 2) && cell.getRowIndex() > 0) {
                        HyTechList.add(cell.getNumericCellValue());
                    }
                }
                System.out.println("");
            }

            //Makes an arraylist with the info needed from FY20 File
            FileInputStream FY20File = new FileInputStream(new File(fy20File));
            XSSFWorkbook FY20Wbk = new XSSFWorkbook(FY20File);
            XSSFSheet FY20Sheet = FY20Wbk.getSheetAt(1);
            Iterator<Row> rowIterator2 = FY20Sheet.iterator();
            while (rowIterator2.hasNext()) {
                Row row2 = rowIterator2.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator2 = row2.cellIterator();
                while (cellIterator2.hasNext()) {
                    Cell cell2 = cellIterator2.next();
                    if ((cell2.getColumnIndex() == 4) && cell2.getRowIndex() > 0) {
                        FY20List.add(cell2.getNumericCellValue());
                    }
                }
                System.out.println("");
            }

            //Makes an arraylist with other info needed from FY20 File
            Iterator<Row> rowIterator3 = FY20Sheet.iterator();
            while (rowIterator3.hasNext()) {
                Row row3 = rowIterator3.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator3 = row3.cellIterator();
                while (cellIterator3.hasNext()) {
                    Cell cell3 = cellIterator3.next();
                    if ((cell3.getColumnIndex() == 3) && cell3.getRowIndex() >= 0) {
                        FY20DescList.add(cell3.getStringCellValue());
                    }
                }
                System.out.println("");
            }
            HyTechFile.close();
            FY20File.close();

            //Create new file to show matches
            XSSFWorkbook confwkbk = new XSSFWorkbook();
            XSSFSheet confSheet = confwkbk.createSheet("Confirmation");
            Map<Integer, Object[]> data = new TreeMap<Integer, Object[]>();
            data.put(1, new Object[]{"Location in FY20 Sheet", "Amount", "Message", "Location in HyTech Sheet", "Item Desc"});
            for (Double i: FY20List) {
                for (Double j: HyTechList) {
                    counter4++;
                    if ((i == -j)) {
                        if (counter5 == 0) {
                            data.put(++counter3, new Object[]{counter3, i, "Match Found", counter4, FY20DescList.get(counter3 - 1)});
                            counter5++;
                        } else {
                            data.put(counter3, new Object[]{counter3, i, "Multiple Matches Found!", counter4, FY20DescList.get(counter3 - 1)});
                        }
                        yeet = true;
                    }
                }
                if (yeet) {
                    yeet = false;
                } else {
                    if (i < 0) {
                        data.put(++counter3, new Object[]{counter3, i, "Match Not Found!", -1, FY20DescList.get(counter3 - 1)});
                    } else {
                        counter3++;
                    }
                }
                counter4 = 1;
                counter5 = 0;
            }
            Set<Integer> keyset = data.keySet();
            int rownum = 0;
            for (Integer key : keyset) {
                if (key <= FY20List.size()) {
                    Row row = confSheet.createRow(rownum++);
                    Object[] objArr = data.get(key);
                    int cellnum = 0;
                    for (Object obj : objArr) {
                        Cell cell = row.createCell(cellnum++);
                        if (obj instanceof String)
                            cell.setCellValue((String) obj);
                        else if (obj instanceof Integer)
                            cell.setCellValue((Integer) obj);
                        else if (obj instanceof Double)
                            cell.setCellValue((Double) obj);
                    }
                }
            }

            //Write the file into the location specified
            FileOutputStream out = new FileOutputStream(new File(confFile + "confirmation.xlsx"));
            confwkbk.write(out);
            out.close();
            System.out.println("confirmation.xlsx written successfully on disk.");
        }
        catch (Exception e) {
            System.out.println("Make sure the files are in the correct locations and have the correct names");
            e.printStackTrace();
        }
    }
}

/*
RRRRRR              OOOOOOOOOOOOOOO     HHH         HHH
RRR   RRR           OOOOOOOOOOOOOOO     HHH         HHH
RRR      RRR        OOO         OOO     HHH         HHH
RRR      RRR        OOO         OOO     HHH         HHH
RRR   RRR           OOO         OOO     HHH         HHH
RRRRRR              OOO         OOO     HHHHHHHHHHHHHHH
RRR  RRR            OOO         OOO     HHH         HHH
RRR   RRR           OOO         OOO     HHH         HHH
RRR    RRR          OOO         OOO     HHH         HHH
RRR     RRR         OOOOOOOOOOOOOOO     HHH         HHH
RRR      RRR        OOOOOOOOOOOOOOO     HHH         HHH
 */
