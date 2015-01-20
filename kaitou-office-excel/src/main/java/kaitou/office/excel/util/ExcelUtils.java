package kaitou.office.excel.util;

import kaitou.office.excel.common.SysCode;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;

import java.io.*;
import java.text.DecimalFormat;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * excel操作工具类.
 * User: 赵立伟
 * Date: 2015/1/10
 * Time: 23:57
 */
public abstract class ExcelUtils {
    /**
     * 获取工作单元
     *
     * @param file excel文件
     * @return 工作单元
     */
    public static Sheet getSheet(File file) throws IOException, InvalidFormatException {
        return create(file).getSheetAt(0);
    }

    /**
     * 复制工作单元
     *
     * @param newSheet  新单元
     * @param sheetName 新单元名
     * @param wb        workbook
     */
    public static void copy(Sheet newSheet, String sheetName, HSSFWorkbook wb) throws IOException, InvalidFormatException {
        Sheet sheetCreate = wb.createSheet(sheetName);
        MergerRegion(sheetCreate, newSheet);
        int firstRow = newSheet.getFirstRowNum();
        int lastRow = newSheet.getLastRowNum();
        for (int i = firstRow; i <= lastRow; i++) {
            // 创建新建excel Sheet的行
            Row rowCreate = sheetCreate.createRow(i);
            // 取得源有excel Sheet的行
            Row row = newSheet.getRow(i);
            // 单元格式样
            int firstCell = row.getFirstCellNum();
            int lastCell = row.getLastCellNum();
            for (int j = firstCell; j < lastCell; j++) {
                // 自动适应列宽 貌似不起作用
                rowCreate.createCell(j);
                String strVal = removeInternalBlank(row.getCell(j).getStringCellValue());
                rowCreate.getCell(j).setCellValue(strVal);
            }
        }
    }

    /**
     * 复制合并单元格
     *
     * @param sheetCreate 新工作单元
     * @param sheet       工作单元
     */
    private static void MergerRegion(Sheet sheetCreate, Sheet sheet) {
        int sheetMergerCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergerCount; i++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            sheetCreate.addMergedRegion(mergedRegion);
        }

    }

    /**
     * 去除字符串内部空格
     *
     * @param s 字符串
     * @return 处理后的串
     */
    private static String removeInternalBlank(String s) {
        Pattern p = Pattern.compile("\\s*|\t|\r|\n");
        Matcher m = p.matcher(s);
        char str[] = s.toCharArray();
        StringBuffer sb = new StringBuffer();
        for (int i = 0; i < str.length; i++) {
            if (str[i] == ' ') {
                sb.append(' ');
            } else {
                break;
            }
        }
        String after = m.replaceAll("");
        return sb.toString() + after;
    }

    /**
     * 根据不同版本excel文件创建工作簿
     *
     * @param file 文件
     * @return 工作簿
     */
    public static Workbook create(File file) throws IOException, InvalidFormatException {
        InputStream is = null;
        try {
            is = new FileInputStream(file);
            if (!is.markSupported()) {
                is = new PushbackInputStream(is, 8);
            }
            if (POIFSFileSystem.hasPOIFSHeader(is)) {
                return new HSSFWorkbook(is);
            }
            if (POIXMLDocument.hasOOXMLHeader(is)) {
                return new XSSFWorkbook(OPCPackage.open(is));
            }
            throw new IllegalArgumentException("你的excel版本目前poi解析不了");
        } finally {
            if (is != null) {
                is.close();
            }
        }
    }


    /**
     * 获取合并单元格数据
     *
     * @param sheet
     * @param x        横坐标范围
     * @param y        纵坐标范围
     * @param cellType 数据类型
     * @return 单元格数据
     */
    public static String getMergedRegions(Sheet sheet, int[] x, int[] y, SysCode.CellType cellType) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();
            if (x[0] == firstRow && x[1] == lastRow) {
                if (y[0] >= firstColumn && y[1] <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);

                    return getStringCellValue(fCell, cellType);
                }
            }
        }
        return "";
    }


    /**
     * 将单元格数据转换成字符串
     *
     * @param cell 单元格
     * @param type 类型
     * @return 字符串
     */
    private static String getStringCellValue(Cell cell, SysCode.CellType type) {
        DecimalFormat df = new DecimalFormat("0");
        try {
            switch (type.getValue()) {
                case SysCode.CELL_TYPE_STRING:
                    return cell.getStringCellValue();
                case SysCode.CELL_TYPE_DATE:
                    return new DateTime(cell.getDateCellValue().getTime()).toString("yyyy/MM/dd");
                case SysCode.CELL_TYPE_NUMERIC:
                    return df.format(cell.getNumericCellValue());
                default:
                    return "";
            }
        } catch (Exception e) {
            if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
                return df.format(cell.getNumericCellValue());
            }
            return "";
        }
    }


    /**
     * 加入到excel文件
     *
     * @param file      excel文件
     * @param sheetName 工作单元名
     * @param datas     新数据
     * @return 总行数
     */
    public static int add2Sheet(File file, String sheetName, List<Object[]> datas) throws IOException, InvalidFormatException {
        Workbook workbook = create(file);
        Sheet sheet = workbook.getSheet(sheetName);
        int allRow = sheet.getLastRowNum();
        Row row;
        for (Object[] data : datas) {
            if (data == null || data.length <= 0) {
                continue;
            }
            row = sheet.createRow(++allRow);
            for (int j = 0; j < data.length; j++) {
                row.createCell(j).setCellValue(String.valueOf(data[j] == null ? "" : data[j]));
            }
        }
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(file);
            workbook.write(out);
            out.flush();
        } finally {
            if (out != null) {
                out.close();
            }
        }
        return allRow;
    }

    /**
     * 获取最后一行单元格数据
     *
     * @param file        excel文件
     * @param sheetName   工作单元名
     * @param columnIndex 列序号
     * @param type        类型
     * @return 单元格数据
     */
    public static String getLastRowCellStrValue(File file, String sheetName, int columnIndex, SysCode.CellType type) throws IOException, InvalidFormatException {
        Workbook workbook = create(file);
        Sheet sheet = workbook.getSheet(sheetName);
        int lastRowNum = sheet.getLastRowNum();
        if (lastRowNum < 1) {
            return "";
        }
        Row row = sheet.getRow(lastRowNum);
        return getStringCellValue(row.getCell(columnIndex), type);
    }
}
