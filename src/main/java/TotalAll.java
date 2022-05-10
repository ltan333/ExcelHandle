import java.util.Date;

public class TotalAll {
    private Date date;
    private double amount;

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public double getAmount() {
        return amount;
    }

    public void setAmount(double amount) {
        this.amount = amount;
    }

    public TotalAll() {
    }

    public TotalAll(Date date, double amount) {
        this.date = date;
        this.amount = amount;
    }
}
