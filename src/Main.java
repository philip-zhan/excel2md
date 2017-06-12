import com.beust.jcommander.JCommander;
import com.beust.jcommander.Parameter;
import com.beust.jcommander.ParameterException;
import com.beust.jcommander.converters.IParameterSplitter;
import javafx.util.Pair;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.*;

public class Main {

//    @Parameter(names = {"--output", "-o"},
//            description = "输出的Markdown文件名（多个文件用空格隔开）",
//            variableArity = true)
//    List<String> outputFileNames;

    public static class SpaceSplitter implements IParameterSplitter {
        public List<String> split(String value) {
            return Arrays.asList(value.split(" "));
        }
    }

    // Main parameter
    @Parameter(description = "需要转换的Excel文件名（多个文件用空格隔开）"
            , required = true
    )
    static List<String> inputFileNames;

    @Parameter(names = {"--sheet", "-s"},
            description = "需要转换Excel文件里的第几张表（从0开始，多张表用逗号隔开，默认全转）",
            required = true,
            splitter = SpaceSplitter.class
    )
    static List<String> sheetsList = new ArrayList<String>();

    @Parameter(names = {"--minify", "-m"},
            description = "不填充多余空格使列对齐"
    )
    static boolean minify = true;

    @Parameter(names = {"--help", "-help"},
            description = "使用说明",
            help = true
    )
    static boolean help = false;


    public static void main(String[] args) {
        Main main = new Main();
        JCommander jCommander = JCommander.newBuilder().addObject(main).build();
        jCommander.setProgramName("xlsx2md");
        try {
            jCommander.parse(args);
            if (main.help){
                jCommander.usage();
                return;
            }
            main.run();
        }
        catch(ParameterException ex){
            jCommander.usage();
        }
    }

    public void run() {
        //TODO: Exception in thread "main" java.lang.NullPointerException at Main.run(Main.java:70)
        if (inputFileNames.size() != sheetsList.size()){
            //TODO: warning message
            return;
        }
        for (int i = 0; i < inputFileNames.size(); i++){
            String inputFileName = inputFileNames.get(i);
            String sheets = sheetsList.get(i);

            File inputFile = readInputFile(inputFileName);
            if (inputFile == null) {
                return;
            }
            parseXls(inputFile, sheets);
        }
    }

    public static File readInputFile(String fileName) {
        if (fileName.isEmpty()) {
            System.out.println("请指定需要转换的文件");
            return null;
        } else if (!(fileName.endsWith(".xls") || fileName.endsWith(".xlsx"))) {
            System.out.println("只能转换xls和xlsx文件");
            return null;
        }
        File inputFile = new File(fileName);
        if (!inputFile.exists()) {
            System.out.println("文件不存在");
            return null;
        } else if (!inputFile.canRead()) {
            System.out.println("无法读取文件");
            return null;
        }
        return inputFile;
    }

    public static BitSet getSheets(Workbook workbook, String sheets) {
        //System.out.println(workbook.getNumberOfSheets());
        BitSet sheetsBitSet = new BitSet();
        //new BitSet(workbook.getNumberOfSheets() - 1);
        // convert all sheets in the workbook
        if (sheets.toLowerCase().contains("all")){
            // set all sheet bits
            sheetsBitSet.set(0, workbook.getNumberOfSheets());
            System.out.println(sheetsBitSet);
            return sheetsBitSet;
        }
        String[] sheetsArray = sheets.split(",");
        for (String sheet : sheetsArray) {
            try {
                sheetsBitSet.set(Integer.parseInt(sheet));
            }catch (NumberFormatException ex) {
                //TODO message
                System.out.println();
                return null;
            }catch (IndexOutOfBoundsException ex) {
                //TODO message
                System.out.println();
                return null;
            }
        }
        return sheetsBitSet;
    }

    public static void parseXls(File inputFile, String sheets) {
        Workbook workbook;
        try {
            workbook = WorkbookFactory.create(inputFile);
        }catch (InvalidFormatException | IOException ex) {
            System.out.println("无法打开文件");
            return;
        }

        BitSet sheetsBitSet = getSheets(workbook, sheets);
        if (sheetsBitSet == null) {
            return;
        }
        // TODO: convert multiple sheets or specify sheet to convert
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
                    for (int j = firstRowIndex; j < rowBitSet.length(); j++) {
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
                    //TODO message
                }
            }
        }
    }

    public static Map.Entry<List<Integer>, BitSet> getDimensions(Sheet sheet) {
        List<Integer> columnWidthList = new ArrayList<>();
        BitSet rowBitSet = new BitSet();
        // getLastRowNum() is 0 based, thus the '<='
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            // if not an empty row, sweep columns
            if (row != null){
//                System.out.println("==================");
//                System.out.println(row.getLastCellNum());
//                System.out.println("------------------");
                int lastCellNum = row.getLastCellNum();
                // iterate through cells
                for (int j = 0; j <= lastCellNum; j++) {
                    //System.out.println(j);
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
        return new AbstractMap.SimpleEntry<>(columnWidthList.subList(0, indexOfLastNonZero + 1), rowBitSet);
    }

    public static void buildTableHeader(Row row, PrintWriter writer, List<Integer> columnWidthList) {
        StringBuilder row1 = new StringBuilder("|");
        StringBuilder row2 = new StringBuilder("|");
        for (int i = 0; i < columnWidthList.size(); i++) {
            //System.out.println(i);
            //System.out.println(columnWidthList);
            String cellContent = getCellContent(row, i, columnWidthList.get(i));
            // build the separator row with each cell having the same width as the header
            String separator = new String(new char[cellContent.length() + 2]).replace("\0", "-");
            row1.append(" ").append(cellContent).append(" |");
            row2.append(separator).append("|");
        }
        writer.println(row1.toString());
        writer.println(row2.toString());
    }

    public static String getCellContent(Row row, int i, int columnWidth) {
        String minifiedCellContent = getMinifiedCellContent(row, i);
        if (minify) {
            return minifiedCellContent;
        }else{
            int additionalSpaces = columnWidth - minifiedCellContent.length();
            if (additionalSpaces < 0){
                additionalSpaces = 0;
            }
            return minifiedCellContent + new String(new char[additionalSpaces]).replace("\0", " ");
        }
    }

    public static String getMinifiedCellContent(Row row, int i) {
        DataFormatter formatter = new DataFormatter();
        //System.out.println(i);
        //System.out.println(row);
        Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        String lineSeparatorRemoved = formatter.formatCellValue(cell).replaceAll(System.lineSeparator(), "");
        String pipeCharacterEscaped = lineSeparatorRemoved.replaceAll("\\|", "\\\\|");
        return pipeCharacterEscaped;
    }

}
