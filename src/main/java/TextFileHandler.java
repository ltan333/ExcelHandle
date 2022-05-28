import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Hashtable;
import java.util.Scanner;

public class TextFileHandler {

    public static File file = new File("Payment_Calculation.txt");

    public static Scanner readData(){
        try {
            FileInputStream fileInputStream = new FileInputStream(file);
            Scanner scan = new Scanner(fileInputStream);
            return scan;

        } catch (FileNotFoundException e) {
            System.out.println("Can't read Payment_Calculation.txt");
            e.printStackTrace();
        }
        return null;
    }

    public static Hashtable defineHowToCalculate(String employeeName){
        Hashtable<Integer,String> ht = new Hashtable<>();
        Scanner scan = readData();
        while (scan.hasNextLine()){
            String name = scan.nextLine();
            if(name.strip().toLowerCase().contains(employeeName.toLowerCase()) && name.startsWith("Name")){
                for (int i = 0; i < 7; i++) {
                    String s = scan.nextLine().strip();
                    try{
                        ht.put(i+1,s.split("=")[1]);
                    }catch (ArrayIndexOutOfBoundsException e){
                        ht.put(i+1,"");
                    }
                }
                return ht;
            }
        }
        for (int i = 0; i < 7; i++) {
            ht.put(i+1,"");
        }
        return ht;
    }


}
