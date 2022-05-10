import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;

public class Employee {
    private String name;
    private ArrayList<SalaryOfDate> allSalaryEachDay = new ArrayList<>();


    public Employee(String name) {
        this.name = name;
    }

    public Employee() {
    }

    public void addADaySalary(SalaryOfDate s){
        allSalaryEachDay.add(s);
    }

    public ArrayList<SalaryOfDate> getAllSalaryEachDay() {
        return allSalaryEachDay;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public SalaryOfDate getSalaryOfDate(Date date){
        for (SalaryOfDate s : allSalaryEachDay){
            if(s.getDate().compareTo(date)==0){
                return s;
            }
        }
        return new SalaryOfDate();
    }

    public double[] calculateSalaryOfEmployeeTwoWeek1(Date date){
        double[] list = new double[3];
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        int totalSalary = 0;
        int totalTip = 0;
        int totalCash = 0;
        for(int i = 0; i<15;i++){
            for(SalaryOfDate s : allSalaryEachDay){
                if(s.getDate().compareTo(calendar.getTime())==0){
                    totalSalary+=s.getSalary();
                    totalTip+=s.getTip();
                    totalCash+=s.getCash();
                }
            }
            calendar.add(Calendar.DAY_OF_MONTH,1);
        }
        list[0] = totalSalary;
        list[1] = totalTip;
        list[2] = totalCash;
        return list;
    }

    public double[] calculateSalaryOfEmployeeTwoWeek2(Date date, int numDateInMonth){
        double[] list = new double[3];
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        int totalSalary = 0;
        int totalTip = 0;
        int totalCash = 0;
        calendar.add(Calendar.DAY_OF_MONTH,15);
        for(int i = 0; i<numDateInMonth-15;i++){
            for(SalaryOfDate s : allSalaryEachDay){
                if(s.getDate().compareTo(calendar.getTime())==0){
                    totalSalary+=s.getSalary();
                    totalTip+=s.getTip();
                    totalCash+=s.getCash();
                }
            }
            calendar.add(Calendar.DAY_OF_MONTH,1);
        }
        list[0] = totalSalary;
        list[1] = totalTip;
        list[2] = totalCash;
        return list;
    }
    public void show(){
        System.out.println(this.getName()+ ": ");
        for (SalaryOfDate s : allSalaryEachDay){
            System.out.println(s.getDateString()+" "+s.getSalary()+" "+s.getTip()+" "+s.getCash());
        }
    }
}
