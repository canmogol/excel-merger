package tr.edu.hacettepe;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import picocli.CommandLine;
import picocli.CommandLine.Command;
import picocli.CommandLine.Parameters;
import picocli.jansi.graalvm.AnsiConsole;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.Callable;

import static java.nio.file.Files.list;
import static java.util.Objects.isNull;

@Command(name = "merger", mixinStandardHelpOptions = true,
        version = "merger 1.0",
        description = "Merges the excel files in the provided folder.")
public class ExcelMerger implements Callable<Integer> {

    @Parameters(index = "0", description = "The directory with excel files ('.xls' files).")
    private File folder;

    private final Set<String> componentNames = new LinkedHashSet<>();
    private final Map<String, Map<String, String>> results = new LinkedHashMap<>();
    private final DecimalFormat decimalFormat = new DecimalFormat("#.#########");

    public static void main(String[] args) throws IOException {
        int exitCode;
        try (AnsiConsole ansi = AnsiConsole.windowsInstall()) {
            exitCode = new CommandLine(new ExcelMerger()).execute(args);
        }
        System.exit(exitCode);
    }

    private void merge(String fileLocation, String folder) throws IOException {
        FileInputStream file = new FileInputStream(fileLocation);
        Workbook workbook = new HSSFWorkbook(file);
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            Row componentNameRow = sheet.getRow(2);
            if (isNull(componentNameRow)) {
                break;
            }
            Cell componentNameCell = componentNameRow.getCell(0);
            if (isNull(componentNameCell)) {
                break;
            }
            String componentName = componentNameCell.getStringCellValue();
            componentName = componentName.replace(",", "-");
            componentNames.add(componentName);

            int rowNumber = 0;
            for (Row row : sheet) {
                rowNumber++;
                if (rowNumber < 6) {
                    continue;
                }
                Cell cell0 = row.getCell(0);
                Cell cell4 = row.getCell(4);
                if (cell0.getCellType().equals(CellType.STRING)
                        && (
                        cell0.getStringCellValue().equals("Created By:") || cell0.getStringCellValue().isBlank()
                )) {
                    break;
                }

                switch (cell0.getCellType()) {
                    case STRING:
                        results.putIfAbsent(cell0.getStringCellValue(), new LinkedHashMap<>());
                        break;
                    default:
                        throw new RuntimeException("Not supported Filename value: '%s' on sheet '%s' in file '%s'"
                                .formatted(cell0.getStringCellValue(), componentName, fileLocation));
                }

                switch (cell4.getCellType()) {
                    case STRING:
                        results.get(cell0.getStringCellValue()).putIfAbsent(componentName, cell4.getStringCellValue());
                        break;
                    case NUMERIC:
                        String value = decimalFormat.format(cell4.getNumericCellValue());
                        results.get(cell0.getStringCellValue()).putIfAbsent(componentName, value);
                        break;
                    default:
                        throw new RuntimeException("Not supported value: '%s' on sheet '%s' in file '%s'"
                                .formatted(cell4.getStringCellValue(), componentName, fileLocation));
                }

            }
        }
        StringBuilder csv = new StringBuilder();
        csv.append("FileName,%s\n".formatted(String.join(",", componentNames)));
        results.forEach((filename, sheetAreaPairs) -> {
            csv.append(filename).append(",");
            for (String sheetName : componentNames) {
                csv.append(sheetAreaPairs.getOrDefault(sheetName, "---")).append(",");
            }
            csv.append("\n");
            try {
                Files.writeString(Paths.get(folder + File.separator + "output.csv"), csv.toString());
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        });
    }

    @Override
    public Integer call() throws Exception {
        final ExcelMerger excelMerger = new ExcelMerger();
        if (isNull(folder) || !folder.isDirectory()) {
            log("Please provide folder location.\nusage: merger /Users/username/excel-files");
            return 74;
        }
        list(folder.toPath())
                .filter(path -> path.toString().toLowerCase().endsWith(".xls"))
                .forEach(path -> {
                    try {
                        excelMerger.merge(path.toString(), folder.toString());
                    } catch (IOException e) {
                        throw new RuntimeException(e);
                    }
                });
        return 0;
    }

    private void log(String message) {
        System.out.println(message);
    }
}