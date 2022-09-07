package de.densk.aufg2;

import java.io.IOException;

public class App {
    private static final String EXCEL_FILE_PATH = String.format(
            "%s\\src\\main\\resources\\table.xlsx",
            System.getProperty("user.dir")
    );

    private static final String CSV_FILE_PATH = String.format(
            "%s\\src\\main\\resources\\table.csv",
            System.getProperty("user.dir")
    );

    public static void main(String[] args) throws IOException {
        // Erste Teilaufgabe
        final Excel3000 excel1 = new Excel3000();

        excel1.setCell("A1", "Hello");
        excel1.setCell("C1", "World");
        excel1.setCell("A3", "Ice Cream");

        System.out.println(excel1.getCellAt(1, 1)); // Hello
        System.out.println(excel1.getCellAt("A1")); // Hello

        System.out.println(excel1.getCellAt("A3")); // Ice Cream
        System.out.println(excel1.getCellAt(3, 1)); // Ice Cream


        // Zweite Teilaufgabe
        final Excel3000 excel2 = new Excel3000();

        excel2.setCell("A1", "4.4");
        excel2.setCell("T3", "6.6");
        excel2.setCell("A2", "= $A1 * 2 + $T3");

        final Excel3000 excelEvaluated = excel2.evaluate();

        System.out.println(excelEvaluated.getCellAt("A1")); // 4.4
        System.out.println(excelEvaluated.getCellAt("T3")); // 6.6
        System.out.println(excelEvaluated.getCellAt("A2")); // 15.4


        // Dritte Teilaufgabe
        excelEvaluated.exportToExcel(EXCEL_FILE_PATH);


        // Vierte Teilaufgabe
        excelEvaluated.exportToCSV(CSV_FILE_PATH);
    }
}
