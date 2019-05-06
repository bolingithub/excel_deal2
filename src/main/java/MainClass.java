import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class MainClass {

    /// 处理文件路径
    private static final String EXCEL_PATH = "/Users/rapaq/Downloads/2019年回款统计_3.xlsx";

    /// 处理之后的文件保存路径
    private static final String DEAL_EXCEL_PATH = "/Users/rapaq/Downloads/";

    public static void main(String[] args) {
        System.out.println("---------- begin... ---------");
        try {
            File excelFile = new File(EXCEL_PATH);
            FileInputStream in = new FileInputStream(excelFile);
            Workbook workbook = new XSSFWorkbook(in);
            int sheetCount = workbook.getNumberOfSheets(); // Sheet的数量
            System.out.println("当前Sheet的数量：" + sheetCount);
            for (int i = 0; i < sheetCount; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                System.out.println("当前 i：" + i + " name:" + sheet.getSheetName());
            }

            // 读取表1数据
            List<List<Cell>> rowList_1 = new ArrayList<List<Cell>>();
            Sheet sheet_1 = workbook.getSheetAt(0);
            readSheetData(sheet_1, rowList_1);

            // 读取表2数据
            List<List<Cell>> rowList_2 = new ArrayList<List<Cell>>();
            Sheet sheet_2 = workbook.getSheetAt(1);
            readSheetData(sheet_2, rowList_2);

            // 读取表3数据
            List<List<Cell>> rowList_3 = new ArrayList<List<Cell>>();
            Sheet sheet_3 = workbook.getSheetAt(2);
            readSheetData(sheet_3, rowList_3);

            // 处理数据
            dealData(rowList_1, rowList_2, rowList_3);

            System.out.println("!!! 处理成功 !!!");
            System.out.println("--------- end... ---------");
        } catch (Exception e) {
            System.out.println("!!! 处理失败 !!! 错误：" + e.getMessage());
            System.out.println("---------- begin... ---------");
        }
    }


    /**
     * 读取表数据
     *
     * @param sheet   表
     * @param rowList 存放在哪里
     */
    private static void readSheetData(Sheet sheet, List<List<Cell>> rowList) {
        System.out.println("当前处理的sheet是：" + sheet.getSheetName());
        System.out.println("当前sheet有：" + sheet.getPhysicalNumberOfRows() + "行");
        for (Row row : sheet) {
            // 一行的数据
            List<Cell> rowData = new ArrayList<Cell>();
            int end = row.getLastCellNum();
            for (int i = 0; i < end; i++) {
                Cell cell = row.getCell(i);
                // 该cell可能为null !!!
                rowData.add(cell);
            }
            rowList.add(rowData);
        }
        System.out.println("获取表数据成功!!!");
    }

    private static void dealData(List<List<Cell>>... lists) throws IOException {
        // 销售订单
        List<List<Cell>> rowList_1 = lists[0];
        // 销售出库单
        List<List<Cell>> rowList_2 = lists[1];
        // 3月份需要处理的数据
        List<List<Cell>> rowList_3 = lists[2];

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("处理结果");

        // 第一行：头
        List<Cell> f1 = rowList_1.get(0);
        List<Cell> f2 = rowList_2.get(0);

        for (int i = 0; i < f1.size(); i++) {
            if (f1.get(i) != null) {
                setCellData(sheet, 0, i + 1, f1.get(i).toString());
            }
        }

        for (int i = 0; i < f2.size(); i++) {
            if (f2.get(i) != null) {
                setCellData(sheet, 0, i + 1 + f1.size(), f2.get(i).toString());
            }
        }

        /// 比较数据，第一行被用掉了，当前准备写入的行
        int currentWriteRow = 1;

        for (List<Cell> cells : rowList_3) {
            Cell firstCell = cells.get(0);
            System.out.println("当前查找的数据是：" + firstCell);
            if (firstCell != null) {
                // 在销售订单中查找
                List<List<Cell>> xiaoshou = findData(firstCell, rowList_1);
                // 在销售出库单中查找
                List<List<Cell>> chuku = findData(firstCell, rowList_2);
                // 写入数据
                currentWriteRow = writeData(firstCell, xiaoshou, chuku, sheet, currentWriteRow, f1.size());
            }
        }

        // 写入到文件中
        Date date = new Date(System.currentTimeMillis());
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String time = sdf.format(date);
        String path = DEAL_EXCEL_PATH + time + ".xlsx";
        FileOutputStream out = new FileOutputStream(path);
        workbook.write(out);
    }

    /// 找到所有数据
    private static List<List<Cell>> findData(Cell findCell, List<List<Cell>> inFindCellData) {
        List<List<Cell>> getAllListCell = new ArrayList<List<Cell>>();
        for (List<Cell> cells : inFindCellData) {
            for (Cell cell : cells) {
                if (findCell != null && cell != null && findCell.toString().equals(cell.toString())) {
                    getAllListCell.add(cells);
                }
            }
        }
        return getAllListCell;
    }

    /// 返回
    private static int writeData(Cell findCell, List<List<Cell>> xiaoshou, List<List<Cell>> chuku, XSSFSheet sheet, int currentRow, int secondColumn) {
        setCellData(sheet, currentRow, 0, findCell.toString());
        if (xiaoshou.isEmpty() && chuku.isEmpty()) {
            System.err.println(" !!!!!!!!!!!!! 未找到 数据：" + findCell);
            setCellData(sheet, currentRow, 1, "未找到数据");
            return currentRow + 1;
        }

        // 写入销售单数据
        int currentTempRow = currentRow;
        for (List<Cell> cellList : xiaoshou) {
            for (int i = 0; i < cellList.size(); i++) {
                if (cellList.get(i) != null) {
                    setCellData(sheet, currentTempRow, i + 1, cellList.get(i).toString());
                }
            }
            currentTempRow += 1;
        }


        // 写入出库单数据
        int currentTempRow_2 = currentRow;
        for (List<Cell> cellList : chuku) {
            for (int i = 0; i < cellList.size(); i++) {
                if (cellList.get(i) != null) {
                    setCellData(sheet, currentTempRow_2, i + 1 + secondColumn, cellList.get(i).toString());
                }
            }
            currentTempRow_2 += 1;
        }

        currentRow = currentTempRow > currentTempRow_2 ? currentTempRow : currentTempRow_2;
        return currentRow;
    }

    /// 设置单元格数据 -- ok
    private static void setCellData(XSSFSheet sheet, int rowIndex, int columnIndex, String value) {
        //System.out.println("单元格：" + rowIndex + "  " + columnIndex + "  :" + value);
        // 从0开始
        XSSFRow row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        XSSFCell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        cell.setCellValue(value);
    }
}
