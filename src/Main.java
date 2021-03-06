import com.beust.jcommander.JCommander;
import com.beust.jcommander.Parameter;
import com.beust.jcommander.ParameterException;
import com.beust.jcommander.converters.IParameterSplitter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.*;

public class Main {

    // custom space splitter class for splitting "--sheet" arguments
    public static class SpaceSplitter implements IParameterSplitter {
        public List<String> split(String value) {
            return Arrays.asList(value.split(" "));
        }
    }

    // Main parameter
    @Parameter(description = "The name of file(s) you want to convert (separate multiple files with a space)",
            required = true
    )
    private static List<String> inputFileNames;
    @Parameter(names = {"--sheet", "-s"},
            description = "[Required] The index of sheet(s) you want to convert (0 indexed, separate multiple indices with a comma)",
            required = true,
            splitter = SpaceSplitter.class,
            order = 0
    )
    private static List<String> sheetsList = new ArrayList<>();
    @Parameter(names = {"--align", "-a"},
            description = "[Optional] Append padding spaces to each cell to align columns",
            order = 1
    )
    private static boolean align = false;
    @Parameter(names = {"--help", "-help"},
            description = "User manual",
            help = true,
            hidden = true
    )
    private static boolean help = false;

    private static FormulaEvaluator evaluator;
    private static DataFormatter formatter = new DataFormatter();

    public static void main(String[] args) {
        Main main = new Main();
        JCommander jCommander = JCommander.newBuilder().addObject(main).build();
        jCommander.setProgramName("excel2md");
        try {
            jCommander.parse(args);
            if (help) {
                jCommander.usage();
                return;
            }
            main.run();
        } catch (NoSuchMethodError | ParameterException ex) {
            jCommander.usage();
        }
    }

    private void run() {
        if (inputFileNames.size() != sheetsList.size()) {
            System.out.println("Error: You must specify the sheet(s) you want to convert after each file name");
            return;
        }
        for (int i = 0; i < inputFileNames.size(); i++) {
            String inputFileName = inputFileNames.get(i);
            String sheets = sheetsList.get(i);

            File inputFile = readInputFile(inputFileName);
            if (inputFile == null) {
                return;
            }
            parseXls(inputFile, sheets);
        }
    }

    private static File readInputFile(String fileName) {
        if (fileName.isEmpty()) {
            System.out.println("Error: Please specify the file(s) you want to convert");
            return null;
        } else if (!(fileName.endsWith(".xls") || fileName.endsWith(".xlsx"))) {
            System.out.println("Error: Invalid file format; can only convert .xls and .xlsx files");
            return null;
        }
        File inputFile = new File(fileName);
        if (!inputFile.exists()) {
            System.out.println("Error: " + fileName + " does not exist");
            return null;
        } else if (!inputFile.canRead()) {
            System.out.println("Error: Could not read the input file " + fileName);
            return null;
        }
        return inputFile;
    }

    private static BitSet getSheets(int sheetsInWorkbook, String sheets) {
        BitSet sheetsBitSet = new BitSet();
        // convert all sheets in the workbook
        if (sheets.toLowerCase().contains("all")){
            // set all sheet bits
            sheetsBitSet.set(0, sheetsInWorkbook);
            return sheetsBitSet;
        }
        String[] sheetsArray = sheets.split(",");
        for (String sheet : sheetsArray) {
            try {
                int sheetIndex = Integer.parseInt(sheet);
                if (sheetIndex < sheetsInWorkbook) {
                    sheetsBitSet.set(sheetIndex);
                } else {
                    // warning
                    System.out.println("Warn: Sheet index out of bound; converted valid sheets if any");
                }
            } catch (NumberFormatException | IndexOutOfBoundsException ex) {
                // exception
                System.out.println("Error: Sheet index must be non-negative integer");
                return null;
            }
        }
        return sheetsBitSet;
    }

    private static void parseXls(File inputFile, String sheets) {
        Workbook workbook;
        try {
            workbook = WorkbookFactory.create(inputFile);
        }catch (InvalidFormatException | IOException ex) {
            System.out.println("Error: Invalid Excel file format");
            return;
        }

        BitSet sheetsBitSet = getSheets(workbook.getNumberOfSheets(), sheets);
        if (sheetsBitSet == null) {
            return;
        }

        evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        // iterate through sheets
        for (int i = 0; i < sheetsBitSet.length(); i++) {
            if (sheetsBitSet.get(i)){
                Sheet sheet = workbook.getSheetAt(i);
                try {
                    // create the output file writer
                    PrintWriter printWriter = new PrintWriter(inputFile.getName()
                            .substring(0, inputFile.getName().lastIndexOf("."))
                            + ("_") + (sheet.getSheetName()) + (".md"));

                    Map.Entry<List<Integer>, BitSet> dimensions = getDimensions(sheet);
                    List<Integer> columnWidthList = dimensions.getKey();
                    BitSet rowBitSet = dimensions.getValue();

                    int firstRowIndex = rowBitSet.nextSetBit(0);
                    buildTableHeader(sheet.getRow(firstRowIndex), printWriter, columnWidthList);
                    // iterate through rows
                    for (int j = firstRowIndex + 1; j < rowBitSet.length(); j++) {
                        // if row is not empty
                        if (rowBitSet.get(j)){
                            Row row = sheet.getRow(j);
                            StringBuilder rowBuilder = new StringBuilder("|");
                            // iterate through cells
                            for (int k = 0; k < columnWidthList.size(); k++) {
                                rowBuilder.append(" ").append(
                                        getCellContent(row, k, columnWidthList.get(k))).append(" |");
                            }
                            printWriter.println(rowBuilder.toString());
                        }
                    }
                    printWriter.close();
                }catch (FileNotFoundException ex) {
                    System.out.println("Error: Could not create the markdown file");
                }
            }
        }
    }

    private static Map.Entry<List<Integer>, BitSet> getDimensions(Sheet sheet) {
        List<Integer> columnWidthList = new ArrayList<>();
        BitSet rowBitSet = new BitSet();
        // getLastRowNum() is 0 based, thus the '<='
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            // if not an empty row, sweep columns
            if (row != null){
                int lastCellNum = row.getLastCellNum();
                // iterate through cells
                for (int j = 0; j <= lastCellNum; j++) {
                    int cellWidth = getMinifiedCellContent(row, j).length();
                    if (j >= columnWidthList.size()) {
                        columnWidthList.add(0);
                    }
                    // if a cell is not empty, the row is not empty
                    if (cellWidth > 0) {
                        rowBitSet.set(i);
                        // update columnWidthList with greater cellWidth values
                        if (cellWidth > columnWidthList.get(j)) {
                            columnWidthList.set(j, cellWidth);
                        }
                    }
                }
            }
        }
        int indexOfLastNonZero = 0;
        for (int i = 0; i < columnWidthList.size(); i++) {
            if (columnWidthList.get(i) != 0) {
                indexOfLastNonZero = i;
            }
        }
        return new AbstractMap.SimpleEntry<>(
                columnWidthList.subList(0, indexOfLastNonZero + 1), rowBitSet);
    }

    private static void buildTableHeader(Row row, PrintWriter writer, List<Integer> columnWidthList) {
        StringBuilder row1 = new StringBuilder("|");
        StringBuilder row2 = new StringBuilder("|");
        for (int i = 0; i < columnWidthList.size(); i++) {
            String cellContent = getCellContent(row, i, columnWidthList.get(i));
            // build the separator row with each cell having the same width as the header
            String separator = new String(new char[cellContent.length() + 2]).replace("\0", "-");
            row1.append(" ").append(cellContent).append(" |");
            row2.append(separator).append("|");
        }
        writer.println(row1.toString());
        writer.println(row2.toString());
    }

    private static String getCellContent(Row row, int i, int columnWidth) {
        String minifiedCellContent = getMinifiedCellContent(row, i);
        if (align) {
            int additionalSpaces = columnWidth - minifiedCellContent.length();
            if (additionalSpaces < 0){
                additionalSpaces = 0;
            }
            return minifiedCellContent + new String(new char[additionalSpaces]).replace("\0", " ");
        }else{
            return minifiedCellContent;
        }
    }

    private static String getMinifiedCellContent(Row row, int i) {
        Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        CellValue cellValue = evaluator.evaluate(cell);
        String lineSeparatorRemoved = formatter.formatCellValue(cell, evaluator)
                .replaceAll(System.lineSeparator(), "");
        return lineSeparatorRemoved.replaceAll("\\|", "\\\\|");
    }

}
