import java.util.Date;

public class GiftSell {
    private Date date;
    private double amount;

    public GiftSell() {
    }

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

    public GiftSell(Date date, double amount) {
        this.date = date;
        this.amount = amount;
    }
}
