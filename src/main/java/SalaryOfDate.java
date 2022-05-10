import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class SalaryOfDate {
    private Date date;
    private double salary;
    private double tip;
    private double cash;
    SimpleDateFormat format = new SimpleDateFormat("MM/dd/yyyy");


    public SalaryOfDate(Date date, double salary, double tip, double cash) {
        this.date = date;
        this.salary = salary;
        this.tip = tip;
        this.cash = cash;
    }

    public SalaryOfDate() {
    }

    public Date getDate() {
        return date;
    }

    public String getDateString(){
        return format.format(date);
    }

    public void setDateString(String date) {
        try {
            this.date = format.parse(date);
        } catch (ParseException e) {
            e.printStackTrace();
            this.date = null;
        }
    }

    public void setDate(Date date){
        this.date = date;
    }
    public double getSalary() {
        return salary;
    }

    public void setSalary(double salary) {
        this.salary = salary;
    }

    public double getTip() {
        return tip;
    }

    public void setTip(double tip) {
        this.tip = tip;
    }

    public double getCash() {
        return cash;
    }

    public void setCash(double cash) {
        this.cash = cash;
    }
}
