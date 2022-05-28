import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.time.YearMonth;
import java.util.*;

public class EmployeeSalaryManager {
    private String folderPath = "";
    private Date dateInFile;
    private final ArrayList<Employee> employees = new ArrayList<>();
    private final ArrayList<GiftSell> giftSellList = new ArrayList<>();
    private final ArrayList<TotalAll> totalAllList= new ArrayList<>();

    Scanner scan = new Scanner(System.in);

    public void getFolderPath() {
        System.out.print("Please enter folder path: ");
        this.folderPath = scan.nextLine();
    }

    public void readData(){
        Calendar calendar = Calendar.getInstance();
        for(int i = 1; i<5; i++){
            String fileName = "Week "+i+".xlsx";

            try {
                InputStream inputStream = new FileInputStream(new File(fileName));
                Workbook workbook = new XSSFWorkbook(inputStream);
                int numberOfSheet = workbook.getNumberOfSheets();

                Sheet sheet = workbook.getSheetAt(0);
                Row firstRow = sheet.getRow(0);
                Cell dateCell = firstRow.getCell(0);
                if(i==1){
                    this.dateInFile = dateCell.getDateCellValue();
                    calendar.setTime(dateInFile);
                }

                for (int sheetIndex = 0; sheetIndex < numberOfSheet-1; sheetIndex++) {
                    sheet = workbook.getSheetAt(sheetIndex);
                    if(sheet != null){
                        int rowIndex = 1;
                        int cellIndex = 1;
                        while (true){
                            Row nameRow = sheet.getRow(rowIndex);
                            Row totalRow = sheet.getRow(rowIndex+18);
                            try{
                                Cell nameCell = nameRow.getCell(cellIndex);
                                if(!nameRow.getCell(cellIndex).getStringCellValue().isEmpty()){
                                    if(!nameCell.getStringCellValue().equalsIgnoreCase("gift sell")){
                                        Cell salaryCell = totalRow.getCell(cellIndex);
                                        Cell tipCell = totalRow.getCell(cellIndex+1);
                                        Cell cashCell = totalRow.getCell(cellIndex+2);
                                        Employee e = new Employee();
                                        e.setName(nameCell.getStringCellValue());
                                        addEmployee(e,calendar.getTime(),salaryCell.getNumericCellValue(),tipCell.getNumericCellValue(),cashCell.getNumericCellValue());

                                    }else if(nameCell.getStringCellValue().equalsIgnoreCase("gift sell")){
                                        Cell gsAmountCell = totalRow.getCell(cellIndex);
                                        GiftSell giftSell = new GiftSell();
                                        giftSell.setDate(calendar.getTime());
                                        giftSell.setAmount(gsAmountCell.getNumericCellValue());
                                        giftSellList.add(giftSell);
                                        break;
                                    }

                                }
                                cellIndex+=3;
                            } catch (NullPointerException e) {
                                cellIndex+=3;
                            }
                        }
                    }
                    totalAllList.add(new TotalAll(calendar.getTime(),calculateTotalAll(calendar.getTime())));
                    calendar.add(Calendar.DAY_OF_MONTH,1);
                }

            } catch (FileNotFoundException e) {
                System.out.println("File not found!");
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public void writeData(){
        try{
            OutputStream outputStream;
            outputStream = new FileOutputStream("Salary.xlsx");
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Salary");
            //Create data in sheet.
            createValue(sheet,workbook);
            //Create file.
            workbook.write(outputStream);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void createValue(Sheet sheet, Workbook wb){
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(dateInFile);
        // Get number of days in month
        int numberDateOfMonth = getNumberOfDayInMonth(calendar.getTime());
        // Set cell style.
        CellStyle cellStyle = wb.createCellStyle();
        CreationHelper createHelper = wb.getCreationHelper();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("MM-dd"));
        // Sort employee for name.
        sort(employees);
        // Start row 0 but cell 2 because 0 = day&month, 1 = day
        int rowIndex = 0;
        int cellIndex = 2;
        while (true){
            Row row = sheet.createRow(rowIndex);
            if(rowIndex==0 || rowIndex == 29){
                for(Employee e : employees){
                    Cell nameCell = row.createCell(cellIndex);
                    nameCell.setCellValue(e.getName().toUpperCase());
                    cellIndex+=3;
                }
                Cell totalCell = row.createCell(cellIndex);
                totalCell.setCellValue("TOTAL DAILY");
                Cell recepCell = row.createCell(cellIndex+1);
                recepCell.setCellValue("Receptionist 1");
                Cell recepCell2 = row.createCell(cellIndex+2);
                recepCell2.setCellValue("Receptionist 2");
                cellIndex=2;
            }else if ((rowIndex >= 1 && rowIndex <=15) || (rowIndex >= 29 && rowIndex <=29+numberDateOfMonth-15)) {
                Cell dateCell = row.createCell(0);
                Cell dayOfWeekCell = row.createCell(1);
                dateCell.setCellValue(calendar.getTime());
                dateCell.setCellStyle(cellStyle);
                dayOfWeekCell.setCellValue(new SimpleDateFormat("EEE",Locale.getDefault()).format(calendar.getTime()));
                //Add value
                for(Employee e : employees){
                    Cell salaryCell = row.createCell(cellIndex);
                    salaryCell.setCellValue(e.getSalaryOfDate(calendar.getTime()).getSalary());
                    Cell tipCell = row.createCell(cellIndex+1);
                    tipCell.setCellValue(e.getSalaryOfDate(calendar.getTime()).getTip());
                    Cell cashCell = row.createCell(cellIndex+2);
                    cashCell.setCellValue(e.getSalaryOfDate(calendar.getTime()).getCash());
                    cellIndex+=3;
                }
                //Add total value
                Cell totalCell = row.createCell(cellIndex);
                totalCell.setCellValue(getATotalAll(calendar.getTime()).getAmount());
                Cell recepCell = row.createCell(cellIndex+1);
                recepCell.setCellValue(0);
                Cell recepCell2 = row.createCell(cellIndex+2);
                recepCell2.setCellValue(0);
                cellIndex=2;
                calendar.add(Calendar.DAY_OF_MONTH,1);
            }else if(rowIndex ==17){
                Cell totalCell = row.createCell(0);
                totalCell.setCellValue("TOTAL");
                for (Employee e:employees){
                    double[] totalList = e.calculateSalaryOfEmployeeTwoWeek1(dateInFile);
                    Cell totalSalary = row.createCell(cellIndex);
                    totalSalary.setCellValue(totalList[0]);
                    Cell totalTip = row.createCell(cellIndex+1);
                    totalTip.setCellValue(totalList[1]);
                    Cell totalCash = row.createCell(cellIndex+2);
                    totalCash.setCellValue(totalList[2]);
                    cellIndex+=3;
                }
                Cell totalAll = row.createCell(cellIndex);
                totalAll.setCellValue(calculateTotalAllTwoWeek1(dateInFile));
                cellIndex=2;
            }else if(rowIndex == 18){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("50%");
                for(Employee e:employees){
                    Cell c = row.createCell(cellIndex,CellType.FORMULA);
                    String colName = CellReference.convertNumToColString(c.getColumnIndex()) + (c.getRowIndex());
                    c.setCellFormula(colName+"*50/100");
                    cellIndex+=3;
                }
                Cell c = row.createCell(cellIndex,CellType.FORMULA);
                String colName = CellReference.convertNumToColString(c.getColumnIndex()) + (c.getRowIndex());
                c.setCellFormula(colName+"*50/100");
                cellIndex=2;
            }else if(rowIndex == 19){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("60%");
                for(Employee e:employees){
                    Cell c = row.createCell(cellIndex,CellType.FORMULA);
                    String colName = CellReference.convertNumToColString(c.getColumnIndex()) + (c.getRowIndex()-1);
                    c.setCellFormula(colName+"*60/100");
                    cellIndex+=3;
                }
                Cell c = row.createCell(cellIndex,CellType.FORMULA);
                String colName = CellReference.convertNumToColString(c.getColumnIndex()) + (c.getRowIndex()-1);
                c.setCellFormula(colName+"*60/100");
                cellIndex=2;
            }else if(rowIndex == 20){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("10%");
                for(Employee e:employees){
                    Cell c = row.createCell(cellIndex,CellType.FORMULA);
                    String colName = CellReference.convertNumToColString(c.getColumnIndex()) + (c.getRowIndex()-2);
                    c.setCellFormula(colName+"*10/100");
                    cellIndex+=3;
                }
                Cell c = row.createCell(cellIndex,CellType.FORMULA);
                String colName = CellReference.convertNumToColString(c.getColumnIndex()) + (c.getRowIndex()-2);
                c.setCellFormula(colName+"*10/100");
                cellIndex=2;
            }
            // Check row
            else if(rowIndex == 22){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("CHECK");
                for(Employee e:employees){
                    String value = (String) TextFileHandler.defineHowToCalculate(e.getName()).get(1);
                    if(!value.isEmpty()){
                        Cell c = row.createCell(cellIndex,CellType.FORMULA);
                        c.setCellFormula(findCellLocation(value,c,1));
                    }
                    cellIndex+=3;
                }
                cellIndex=2;
            }
            // Cash Without Tip row
            else if(rowIndex == 23){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("Cash Without Tip");
                for(Employee e:employees){
                    String value = (String) TextFileHandler.defineHowToCalculate(e.getName()).get(2);
                    if(!value.isEmpty()){
                        Cell c = row.createCell(cellIndex,CellType.FORMULA);
                        c.setCellFormula(findCellLocation(value,c,2));
                    }
                    cellIndex+=3;
                }
                cellIndex=2;
            }
            // 60% row
            else if(rowIndex == 24){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("60%");
                for(Employee e:employees){
                    String value = (String) TextFileHandler.defineHowToCalculate(e.getName()).get(3);
                    if(!value.isEmpty()){
                        Cell c = row.createCell(cellIndex,CellType.FORMULA);
                        c.setCellFormula(findCellLocation(value,c,3));
                    }
                    cellIndex+=3;
                }
                cellIndex=2;
            }
            // Check 2 row
            else if(rowIndex == 25){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("Check 2");
                for(Employee e:employees){
                    String value = (String) TextFileHandler.defineHowToCalculate(e.getName()).get(4);
                    if(!value.isEmpty()){
                        Cell c = row.createCell(cellIndex,CellType.FORMULA);
                        c.setCellFormula(findCellLocation(value,c,4));
                    }
                    cellIndex+=3;
                }
                cellIndex=2;
            }
            // 15% row
            else if(rowIndex == 26){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("15%");
                for(Employee e:employees){
                    String value = (String) TextFileHandler.defineHowToCalculate(e.getName()).get(5);
                    if(!value.isEmpty()){
                        Cell c = row.createCell(cellIndex,CellType.FORMULA);
                        c.setCellFormula(findCellLocation(value,c,5));
                    }
                    cellIndex+=3;
                }
                cellIndex=2;
            }
            // Cash Payment row
            else if(rowIndex == 27){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("Cash Payment");
                for(Employee e:employees){
                    String value = (String) TextFileHandler.defineHowToCalculate(e.getName()).get(6);
                    if(!value.isEmpty()){
                        Cell c = row.createCell(cellIndex,CellType.FORMULA);
                        c.setCellFormula(findCellLocation(value,c,6));
                    }
                    cellIndex+=3;
                }
                cellIndex=2;
            }
            //Total payment row
            else if(rowIndex == 28){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("Total payment");
                for(Employee e:employees){
                    String value = (String) TextFileHandler.defineHowToCalculate(e.getName()).get(7);
                    if(!value.isEmpty()){
                        Cell c = row.createCell(cellIndex,CellType.FORMULA);
                        c.setCellFormula(findCellLocation(value,c,7));
                    }
                    cellIndex+=3;
                }
                cellIndex=2;
            }
            else if(rowIndex == 30+numberDateOfMonth-14){
                Cell totalCell = row.createCell(0);
                totalCell.setCellValue("TOTAL");
                for (Employee e:employees){
                    double[] totalList = e.calculateSalaryOfEmployeeTwoWeek2(dateInFile,numberDateOfMonth);
                    Cell totalSalary = row.createCell(cellIndex);
                    totalSalary.setCellValue(totalList[0]);
                    Cell totalTip = row.createCell(cellIndex+1);
                    totalTip.setCellValue(totalList[1]);
                    Cell totalCash = row.createCell(cellIndex+2);
                    totalCash.setCellValue(totalList[2]);
                    cellIndex+=3;
                }
                Cell totalAll = row.createCell(cellIndex);
                totalAll.setCellValue(calculateTotalAllTwoWeek2(dateInFile,numberDateOfMonth));
                cellIndex=2;
            }else if(rowIndex == 30+numberDateOfMonth-13){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("50%");
                for(Employee e:employees){
                    Cell c = row.createCell(cellIndex,CellType.FORMULA);
                    String colName = CellReference.convertNumToColString(c.getColumnIndex()) + (c.getRowIndex());
                    c.setCellFormula(colName+"*50/100");
                    cellIndex+=3;
                }
                Cell c = row.createCell(cellIndex,CellType.FORMULA);
                String colName = CellReference.convertNumToColString(c.getColumnIndex()) + (c.getRowIndex());
                c.setCellFormula(colName+"*50/100");
                cellIndex=2;
            }else if(rowIndex == 30+numberDateOfMonth-12){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("60%");
                for(Employee e:employees){
                    Cell c = row.createCell(cellIndex,CellType.FORMULA);
                    String colName = CellReference.convertNumToColString(c.getColumnIndex()) + (c.getRowIndex()-1);
                    c.setCellFormula(colName+"*60/100");
                    cellIndex+=3;
                }
                Cell c = row.createCell(cellIndex,CellType.FORMULA);
                String colName = CellReference.convertNumToColString(c.getColumnIndex()) + (c.getRowIndex()-1);
                c.setCellFormula(colName+"*60/100");
                cellIndex=2;
            }else if(rowIndex == 30+numberDateOfMonth-11){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("10%");
                for(Employee e:employees){
                    Cell c = row.createCell(cellIndex,CellType.FORMULA);
                    String colName = CellReference.convertNumToColString(c.getColumnIndex()) + (c.getRowIndex()-2);
                    c.setCellFormula(colName+"*10/100");
                    cellIndex+=3;
                }
                Cell c = row.createCell(cellIndex,CellType.FORMULA);
                String colName = CellReference.convertNumToColString(c.getColumnIndex()) + (c.getRowIndex()-2);
                c.setCellFormula(colName+"*10/100");
                cellIndex=2;
            }
            // Check row
            else if(rowIndex == 30+numberDateOfMonth-9){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("CHECK");
                for(Employee e:employees){
                    String value = (String) TextFileHandler.defineHowToCalculate(e.getName()).get(1);
                    if(!value.isEmpty()){
                        Cell c = row.createCell(cellIndex,CellType.FORMULA);
                        c.setCellFormula(findCellLocation(value,c,1));
                    }
                    cellIndex+=3;
                }
                cellIndex=2;
            }
            // Cash Without Tip row
            else if(rowIndex == 30+numberDateOfMonth-8){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("Cash Without Tip");
                for(Employee e:employees){
                    String value = (String) TextFileHandler.defineHowToCalculate(e.getName()).get(2);
                    if(!value.isEmpty()){
                        Cell c = row.createCell(cellIndex,CellType.FORMULA);
                        c.setCellFormula(findCellLocation(value,c,2));
                    }
                    cellIndex+=3;
                }
                cellIndex=2;
            }
            // 60% row
            else if(rowIndex == 30+numberDateOfMonth-7){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("60%");
                for(Employee e:employees){
                    String value = (String) TextFileHandler.defineHowToCalculate(e.getName()).get(3);
                    if(!value.isEmpty()){
                        Cell c = row.createCell(cellIndex,CellType.FORMULA);
                        c.setCellFormula(findCellLocation(value,c,3));
                    }
                    cellIndex+=3;
                }
                cellIndex=2;
            }
            // Check 2 row
            else if(rowIndex == 30+numberDateOfMonth-6){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("Check 2");
                for(Employee e:employees){
                    String value = (String) TextFileHandler.defineHowToCalculate(e.getName()).get(4);
                    if(!value.isEmpty()){
                        Cell c = row.createCell(cellIndex,CellType.FORMULA);
                        c.setCellFormula(findCellLocation(value,c,4));
                    }
                    cellIndex+=3;
                }
                cellIndex=2;
            }
            // 15% row
            else if(rowIndex == 30+numberDateOfMonth-5){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("15%");
                for(Employee e:employees){
                    String value = (String) TextFileHandler.defineHowToCalculate(e.getName()).get(5);
                    if(!value.isEmpty()){
                        Cell c = row.createCell(cellIndex,CellType.FORMULA);
                        c.setCellFormula(findCellLocation(value,c,5));
                    }
                    cellIndex+=3;
                }
                cellIndex=2;
            }
            // Cash Payment row
            else if(rowIndex == 30+numberDateOfMonth-4){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("Cash Payment");
                for(Employee e:employees){
                    String value = (String) TextFileHandler.defineHowToCalculate(e.getName()).get(6);
                    if(!value.isEmpty()){
                        Cell c = row.createCell(cellIndex,CellType.FORMULA);
                        c.setCellFormula(findCellLocation(value,c,6));
                    }
                    cellIndex+=3;
                }
                cellIndex=2;
            }
            //Total payment row
            else if(rowIndex == 30+numberDateOfMonth-3){
                Cell percenCell = row.createCell(0);
                percenCell.setCellValue("Total payment");
                for(Employee e:employees){
                    String value = (String) TextFileHandler.defineHowToCalculate(e.getName()).get(7);
                    if(!value.isEmpty()){
                        Cell c = row.createCell(cellIndex,CellType.FORMULA);
                        c.setCellFormula(findCellLocation(value,c,7));
                    }
                    cellIndex+=3;
                }
                cellIndex=2;
                break;
            }

            rowIndex++;
        }
    }


    public void addEmployee(Employee newEmployee, Date date, double salary, double tip, double cash){
        boolean exist = false;
        for(Employee e: employees){
            if(e.getName().strip().equalsIgnoreCase(newEmployee.getName().strip())){
                e.addADaySalary(new SalaryOfDate(date,salary,tip,cash));
                exist = true;
            }
        }
        if(!exist){
            newEmployee.addADaySalary(new SalaryOfDate(date,salary,tip,cash));
            employees.add(newEmployee);
        }else {

        }
    }

    public double calculateTotalAll(Date date){
        double total = 0;
        for (Employee e:employees){
            for(SalaryOfDate salaryOfDate:e.getAllSalaryEachDay()){
                if(salaryOfDate.getDate().compareTo(date)==0){
                    total+=salaryOfDate.getSalary()+salaryOfDate.getTip()+salaryOfDate.getCash();
                }
            }
        }
        for(GiftSell gs: giftSellList){
            if(gs.getDate().compareTo(date)==0){
                total+=gs.getAmount();
            }
        }
        return total;
    }


    public int getNumberOfDayInMonth(Date date){
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        YearMonth yearMonthObject = YearMonth.of(calendar.get(Calendar.YEAR),calendar.get(Calendar.MONTH)+1);
        int daysInMonth = yearMonthObject.lengthOfMonth();
        return daysInMonth;
    }

    private void sort(ArrayList<Employee> list)
    {
        list.sort((o1, o2) -> o1.getName().compareTo(o2.getName()));
    }

    private TotalAll getATotalAll(Date date){
        for(TotalAll t : totalAllList){
            if(t.getDate().compareTo(date)==0){
                return t;
            }
        }
        return new TotalAll();
    }

    private double calculateTotalAllTwoWeek1(Date date){
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        int total = 0;
        for(int i = 0; i<15;i++){
            for(TotalAll t : totalAllList){
                if(t.getDate().compareTo(calendar.getTime())==0){
                    total+=t.getAmount();
                }
            }
            calendar.add(Calendar.DAY_OF_MONTH,1);
        }
        return total;
    }

    private double calculateTotalAllTwoWeek2(Date date, int numberOfDateInMonth){
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        calendar.add(Calendar.DAY_OF_MONTH,15);
        int total = 0;
        for(int i = 0; i<numberOfDateInMonth-15;i++){
            for(TotalAll t : totalAllList){
                if(t.getDate().compareTo(calendar.getTime())==0){
                    total+=t.getAmount();
                }
            }
            calendar.add(Calendar.DAY_OF_MONTH,1);
        }
        return total;
    }

    private String findCellLocation(String text, Cell cell, int currentIndex){
        if(text.equalsIgnoreCase("0"))
            return "0";

        char[] chars = text.replace(" ","").toCharArray();
        String temp = "";
        ArrayList<String> list = new ArrayList<>();
        int i = 0;
        for (char c :chars) {
            if(c != '+' && c != '-' && c != '*' && c != '/'){
                temp+=c;
                if(i++ == chars.length-1){
                    list.add(temp);
                    temp="";
                }
                continue;
            }
            list.add(temp);
            list.add(c+"");
            temp="";
            i++;
        }
        for(String t: list){
            if(t.length()>1)
                temp+=findLocation(t,currentIndex,cell);
            else
                temp+=t;
        }

        return temp;
    }

    private String findLocation(String text, int currentIndex, Cell cell){
        try{
            double num = Double.parseDouble(text);
            return num+"";
        }catch (NumberFormatException e){
            Hashtable<String,String> ht = new Hashtable<>();
            ht.put("Total_100",CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex()-3-currentIndex));
            ht.put("Total_50",CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex()-2-currentIndex));
            ht.put("Total_60",CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex()-1-currentIndex));
            ht.put("Total_10",CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex()-currentIndex));

            ht.put("Total_Tip_100",CellReference.convertNumToColString(cell.getColumnIndex()+1) + (cell.getRowIndex()-3-currentIndex));
            ht.put("Total_Tip_50",CellReference.convertNumToColString(cell.getColumnIndex()+1) + (cell.getRowIndex()-3-currentIndex) + "*50");
            ht.put("Total_Tip_60",CellReference.convertNumToColString(cell.getColumnIndex()+1) + (cell.getRowIndex()-3-currentIndex) + "*60");
            ht.put("Total_Tip_10",CellReference.convertNumToColString(cell.getColumnIndex()+1) + (cell.getRowIndex()-3-currentIndex) + "*10");

            ht.put("Check", CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex()+2-currentIndex));
            ht.put("Cash_Without_Tip", CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex()+3-currentIndex));
            ht.put("60_Percent", CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex()+4-currentIndex));
            ht.put("Check_2", CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex()+5-currentIndex));
            ht.put("15_Percent", CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex()+6-currentIndex));
            ht.put("Cash_Payment", CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex()+7-currentIndex));
            ht.put("Total_Payment", CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex()+8-currentIndex));

            for(String t: ht.keySet()){
                if(t.equalsIgnoreCase(text)){
                    return ht.get(t);
                }
            }
        }
        return "";
    }
}
