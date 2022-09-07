package de.densk.aufg2;

import com.google.common.collect.HashBasedTable;
import com.google.common.collect.Table;
import net.objecthunter.exp4j.ExpressionBuilder;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Excel3000 {
    private final Table<Integer, Integer, String> table;
    private final String regExCell = "^([a-z]|[A-Z])([1-9]|([1-9][0-9])|100)$";
    private final String regExLetter = "([a-z]|[A-Z])";
    private final String regExCellVariable = "\\$([a-z]|[A-Z])([1-9]|([1-9][0-9])|100)";

    public Excel3000() {
        table = HashBasedTable.create();
    }

    public Excel3000(Table<Integer, Integer, String> table) {
        this.table = table;
    }

    private int letterToColumn(char letter) {
        int col = Character.isUpperCase(letter) ?
                ((int) letter) - ((int) 'A') :
                ((int) letter) - ((int) 'a');

        return col + 1;
    }

    private char columnToLetter(int column) {
        return (char) ((column - 1) + ((int) 'A'));
    }

    public void setCell(String cell, String content) {
        if (cell.matches(regExCell)) {
            if (content != null && !content.isEmpty()) {
                final var col = letterToColumn(cell.charAt(0));
                final var row = Integer.parseInt(cell.replaceAll(regExLetter, ""));

                table.put(row, col, content);
            }
        }
    }

    public String getCellAt(int row, int col) {
        return table.get(row, col);
    }

    public String getCellAt(String cell) {
        String content = null;

        if (cell.matches(regExCell)) {
            final var col = letterToColumn(cell.charAt(0));
            final var row = Integer.parseInt(cell.replaceAll(regExLetter, ""));

            content = getCellAt(row, col);
        }

        return content;
    }

    public Excel3000 evaluate() {
        final Table<Integer, Integer, String> tableEvaluated = HashBasedTable.create();

        table.cellSet().stream().forEach(cell -> {
            String cellValue = cell.getValue();
            final Map<String, String> variables = new HashMap<>();

            if (cellValue.contains("=")) {
                final Pattern pattern = Pattern.compile(regExCellVariable);
                final Matcher matcher = pattern.matcher(cellValue);

                while (matcher.find()) {
                    final var variable = matcher.group();
                    final var value = getCellAt(variable.split("\\$")[1]);

                    variables.put(variable, value);
                }

                if (!variables.isEmpty()) {
                    String evaluatedContent = cellValue;

                    for (var k : variables.keySet()) {
                        final var v = variables.get(k);
                        k = new StringBuilder(k).insert(0, "\\").toString();

                        evaluatedContent = evaluatedContent.replaceAll(k, v);
                    }

                    evaluatedContent = evaluatedContent.replaceAll("=\\s?", "");
                    String calculatedContent = null;

                    try {
                        final double res = new ExpressionBuilder(evaluatedContent).build().evaluate();
                        calculatedContent = res + "";
                    } catch (Exception e) {
                        calculatedContent = "Fehler!";
                    }

                    cellValue = calculatedContent;
                }
            }

            tableEvaluated.put(cell.getRowKey(), cell.getColumnKey(), cellValue);
        });

        return new Excel3000(tableEvaluated);
    }

    public void exportToExcel(String filePath) throws IOException {
        final XSSFWorkbook workbook = new XSSFWorkbook();
        final XSSFSheet sheet = workbook.createSheet();

        table.cellSet().stream().forEach(cellTable -> {
            final Row row = sheet.createRow(cellTable.getRowKey() - 1);
            final Cell cell = row.createCell(cellTable.getColumnKey() - 1);

            cell.setCellValue(cellTable.getValue());
        });

        Path path = Paths.get(filePath);
        workbook.write(Files.newOutputStream(path));
    }

    public void exportToCSV(String filePath) throws IOException {
        final Path path = Paths.get(filePath);

        try (CSVPrinter csvPrinter = new CSVPrinter(Files.newBufferedWriter(path), CSVFormat.DEFAULT)) {
            csvPrinter.printRecord(Arrays.asList("Zeile", "Spalte", "Zelle", "Wert"));

            table.cellSet().stream().forEach(cell -> {
                try {
                    csvPrinter.printRecord(
                            cell.getRowKey(),
                            cell.getColumnKey(),
                            String.format("%s%d", columnToLetter(cell.getColumnKey()), cell.getRowKey()),
                            cell.getValue()
                    );
                } catch (IOException e) {
                }
            });
        }
    }
}
