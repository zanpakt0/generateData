package main.java;


import java.io.IOException;

public class Main {
    public static void main(String... args) throws IOException {
        Human[] humans = Generator.generate(Integer.parseInt(args[0]));
        String fileNameOut = "humans.xls";
        RighterToExcel.righter(humans, fileNameOut);
    }
}
