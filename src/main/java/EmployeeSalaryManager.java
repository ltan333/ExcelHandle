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
    private ArrayList<Employee> employees = new ArrayList<>();
    private ArrayList<GiftSell> giftSellList = new ArrayList<>();
    private ArrayList<TotalAll> totalAllList= new ArrayList<>();

    Scanner scan = new Scanner(System.in);

    public void getFolderPath() {
        System.out.print("Please enter folder path: ");
        this.folderPath = scan.nextLine();
    }

    public void readData(){
        getFolderPath();
        Calendar calendar = Calendar.getInstance();
        for(int i = 1; i<5; i++){
            String fileName = folderPath+"\\Week "+i+".xlsx";
            System.out.println(fileName);
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
        for(Employee e: employees){
            e.show();
        }
    }

    public void writeData(){
        try{
            OutputStream outputStream = new FileOutputStream(new File(folderPath+"\\Salary.xlsx"));
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Salary");
            createValue(sheet,workbook);
            workbook.write(outputStream);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void createValue(Sheet sheet, Workbook wb){
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(dateInFile);
        int numberDateOfMonth = getNumberOfDayInMonth(calendar.getTime());
        CellStyle cellStyle = wb.createCellStyle();
        CreationHelper createHelper = wb.getCreationHelper();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("MM-dd"));
        sort(employees);
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
            }else if(rowIndex == 30+numberDateOfMonth-14){
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
        System.out.println(daysInMonth);
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


}
