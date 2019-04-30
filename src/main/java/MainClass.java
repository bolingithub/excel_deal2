import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

public class MainClass {
    public static void main(String[] args) {
        System.out.println("---------- begin... ---------");
        try {
            File excelFile = new File("/Users/rapaq/Downloads/2019年回款统计_3.xlsx");
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
        // 打印每行数据测试
//        for (List<Cell> list : rowList) {
//            StringBuilder builder = new StringBuilder();
//            for (Cell cell : list) {
//                if (cell == null) {
//                    builder.append(" null ");
//                } else {
//                    builder.append(cell.toString());
//                }
//
//                builder.append("  ");
//            }
//            System.out.println("每行的数据是：" + builder.toString());
//        }
        System.out.println("获取表数据成功!!!");
    }

    private static void dealData(List<List<Cell>>... lists) {
        // 销售订单
        List<List<Cell>> rowList_1 = lists[0];
        // 销售出库单
        List<List<Cell>> rowList_2 = lists[1];
        // 3月份需要处理的数据
        List<List<Cell>> rowList_3 = lists[2];

        for (List<Cell> cells : rowList_3) {
            Cell firstCell = cells.get(0);
            // firstCell 可能为null
            System.out.println("当前查找的数据是：" + firstCell);
            if (firstCell != null) {
                // 在销售订单中查找
                List<Cell> xiaoshou = findData(firstCell, rowList_1);
                // 在销售出库单中查找
                List<Cell> chuku = findData(firstCell, rowList_2);

                writeData(firstCell, xiaoshou, chuku);
            }
        }
    }

    private static List<Cell> findData(Cell findCell, List<List<Cell>> inFindCellData) {
        for (List<Cell> cells : inFindCellData) {
            for (Cell cell : cells) {
                if (findCell != null && cell != null && findCell.toString().equals(cell.toString())) {
                    return cells;
                }
            }
        }
        return null;
    }

    private static void writeData(Cell findCell, List<Cell> xiaoshou, List<Cell> chuku) {
        if (xiaoshou == null && chuku == null) {
            System.err.println(" !!!! 未找到 数据：" + findCell);
            return;
        }
        System.out.println("找到对应的数据是：" + findCell);
        if (xiaoshou != null) {
            printRow(xiaoshou);
        }
        if (chuku != null) {
            printRow(chuku);
        }
    }

    private static void printRow(List<Cell> list) {
        StringBuilder builder = new StringBuilder();
        for (Cell cell : list) {
            if (cell == null) {
                builder.append(" null ");
            } else {
                builder.append(cell.toString());
            }
            builder.append("  ");
        }
        System.out.println("每行的数据是：" + builder.toString());
    }
}
