import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.BitSet;

public class xlsx2md {
    public static void main(String[] args) {
        if (args.length == 0) {
            System.out.println("请指定需要转换的文件");
            return;
        }
        String fileName = args[0];
        File inputFile = readInputFile(fileName);
        if (inputFile == null) {
            return;
        }
        parseXls(inputFile);
    }

    private static File readInputFile(String fileName) {
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

    private static void parseXls(File inputFile) {
        Workbook workbook;
        try {
            PrintWriter writer = new PrintWriter(inputFile.getName().
                    substring(0, inputFile.getName().lastIndexOf(".")).concat(".md"));
            workbook = WorkbookFactory.create(inputFile);
            //TODO: convert multiple sheets or specify sheet to convert
            Sheet sheet = workbook.getSheetAt(1);
            int[] rowsColumns = getDimensions(sheet);
            buildTableHeader(sheet.getRow(0), writer, rowsColumns[1]);
            for (int i = 1; i < rowsColumns[0]; i++) {
                Row row = sheet.getRow(i);
                StringBuilder rowBuilder = new StringBuilder("|");
                for (int j = 0; j < rowsColumns[1]; j++) {
                    rowBuilder.append(" ").append(getCellContent(row, j)).append(" |");
                }
                writer.println(rowBuilder.toString());
            }
            writer.close();
        } catch (IOException | InvalidFormatException e) {
            System.out.println("无法打开文件");
        }
    }

    private static int[] getDimensions(Sheet sheet) {
        int columns = 0;
        int rows = 0;
        BitSet columnBitSet = new BitSet();
        BitSet rowBitSet = new BitSet();
        // getLastRowNum() is 0 based, thus the '<='
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            for (int j = 0; j <= row.getLastCellNum(); j++) {
                if (row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) != null) {
                    columnBitSet.set(j);
                }
            }
            if (columns < columnBitSet.length()) {
                columns = columnBitSet.length();
            }
            // there's at least one set bit (at least one non-empty cell)
            if (columnBitSet.length() > 0) {
                rowBitSet.set(i);
            }
            columnBitSet.clear();
        }
        rows = rowBitSet.length();
        return new int[]{rows, columns};
    }

    private static void buildTableHeader(Row row, PrintWriter writer, int columns) {
        StringBuilder row1 = new StringBuilder("|");
        StringBuilder row2 = new StringBuilder("|");
        for (int i = 0; i < columns; i++) {
            String cellContent = getCellContent(row, i);
            row1.append(" ").append(cellContent).append(" |");
            row2.append(new String(new char[cellContent.length() + 2]).replace("\0", "-")).append("|");
        }
        writer.println(row1.toString());
        writer.println(row2.toString());
    }

    private static String getCellContent(Row row, int i) {
        DataFormatter formatter = new DataFormatter();
        Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        String lineSeparatorRemoved = formatter.formatCellValue(cell).replaceAll(System.lineSeparator(), "");
        String pipeCharacterEscaped = lineSeparatorRemoved.replaceAll("\\|", "\\\\|");
        return pipeCharacterEscaped;
    }

}
