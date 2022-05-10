
import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;

public class Main {

    public static void main(String[] args) throws IOException, ParseException {
        EmployeeSalaryManager e = new EmployeeSalaryManager();
        e.readData();
        e.writeData();
    }

}
