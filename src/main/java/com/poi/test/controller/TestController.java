package com.poi.test.controller;

import org.apache.commons.collections4.iterators.ArrayListIterator;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.ListIterator;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Created with IntelliJ IDEA.
 * User: wb252654
 * Date: 17-9-6
 * Time: 上午10:12
 * To change this template use File | Settings | File Templates.
 */
@Controller
@RequestMapping("test")
public class TestController {

    @RequestMapping("/poi")
    public String poi() {
        System.out.println("welcome");
        return "hello";
    }

    @RequestMapping("upload")
    public void upload(@RequestParam MultipartFile file, HttpServletResponse response) {
        String name = file.getOriginalFilename();
        if (name.endsWith("xls")) {
            parseXlsExcel(file, response);
        } else if (name.endsWith("xlsx")) {
            parseXlsxExcel(file, response);
        }
    }

    //03-07版本
    private void parseXlsExcel(MultipartFile file, HttpServletResponse response) {
        try {
            POIFSFileSystem fs = new POIFSFileSystem(file.getInputStream());
            //得到Excel工作簿对象
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            //取得sheet的数目
            int sheetCount = wb.getNumberOfSheets();
            for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
                //得到Excel工作表对象
                HSSFSheet sheet = wb.getSheetAt(sheetIndex);

                for (int rowIndex = 1; rowIndex < sheet.getLastRowNum(); rowIndex++) {
                    //得到Excel工作表的行
                    HSSFRow row = sheet.getRow(rowIndex);
                    for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
                        //得到Excel工作表指定行的单元格
                        HSSFCell cell = row.getCell((short) cellIndex);
                        HSSFCellStyle cellStyle = cell.getCellStyle();//得到单元格样式
                        System.out.println(getCellStringValue(cell));
                    }
                }
            }
            createXlsExcel(wb, response, file);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    public static void main(String[] args) {
    }

    private String getCellStringValue(HSSFCell cell) {
        String cellValue = "";
        switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_STRING://字符串类型
                cellValue = cell.getStringCellValue();
                if (cellValue.trim().equals("") || cellValue.trim().length() <= 0)
                    cellValue = " ";
                break;
            case HSSFCell.CELL_TYPE_NUMERIC: //数值类型
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case HSSFCell.CELL_TYPE_FORMULA: //公式
                cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case HSSFCell.CELL_TYPE_BLANK:
                cellValue = " ";
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                break;
            case HSSFCell.CELL_TYPE_ERROR:
                break;
            default:
                break;
        }
        return cellValue;
    }

    //07以上版本
    private void parseXlsxExcel(MultipartFile file, HttpServletResponse response) {
        try {
            // 创建 Excel 文件的输入流对象
//            FileInputStream excelFileInputStream = new FileInputStream("D:/employees.xlsx");
// XSSFWorkbook 就代表一个 Excel 文件

// 创建其对象，就打开这个 Excel 文件
            XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());

            if (null != workbook) {
                int count = workbook.getNumberOfSheets();//取得sheet的数目
                for (int sheetIndex = 0; sheetIndex < count; sheetIndex++) {
// 输入流使用后，及时关闭！这是文件流操作中极好的一个习惯！
//            excelFileInputStream.close();
// XSSFSheet 代表 Excel 文件中的一张表格
// 我们通过 getSheetAt(0) 指定表格索引来获取对应表格
// 注意表格索引从 0 开始！
                    XSSFSheet sheet = workbook.getSheetAt(sheetIndex);

                    //选中指定的工作表
//                   sheet.setSelected(true);

                    mergeCell(sheet);

                    // 开始循环表格数据,表格的行索引从 0 开始
// employees.xlsx 第一行是标题行，我们从第二行开始, 对应的行索引是 1
// sheet.getLastRowNum() : 获取当前表格中最后一行数据对应的行索引
                    for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
// XSSFRow 代表一行数据
                        XSSFRow row = sheet.getRow(rowIndex);
                        if (row == null) {
                            continue;
                        }


                        XSSFCell nameCell = row.getCell(0); // 姓名列
                        XSSFCell genderCell = row.getCell(1); // 性别列
                        XSSFCell ageCell = row.getCell(2); // 年龄列
                        XSSFCell weightCell = row.getCell(3); // 体重列
                        XSSFCell salaryCell = row.getCell(4); // 收入列
                        StringBuilder employeeInfoBuilder = new StringBuilder();
                        employeeInfoBuilder.append("员工信息 --> ")
                                .append("姓名 : ").append(nameCell.getStringCellValue())
                                .append(" , 性别 : ").append(genderCell.getStringCellValue())
                                .append(" , 年龄 : ").append(ageCell.getNumericCellValue())
                                .append(" , 体重(千克) : ").append(weightCell.getNumericCellValue())
                                .append(" , 月收入(元) : ").append(salaryCell.getNumericCellValue());
                        System.out.println(employeeInfoBuilder.toString());
                    }
                }
                createExcel(workbook, response);
// 操作完毕后，记得要将打开的 XSSFWorkbook 关闭
                workbook.close();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private XSSFCellStyle setStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();
        style.setBorderBottom(HSSFCellStyle.BORDER_DOTTED);//下边框
        style.setBorderLeft(HSSFCellStyle.BORDER_DOTTED);//左边框
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
        XSSFFont f  = workbook.createFont();
        f.setFontHeightInPoints((short) 11);//字号
        f.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);//加粗
        style.setFont(f);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);//左右居中
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//上下居中
        XSSFDataFormat df = workbook.createDataFormat();
        style.setDataFormat(df.getFormat("0.00%"));//设置单元格数据格式
        //填充和颜色设置
        style.setFillBackgroundColor(HSSFColor.AQUA.index);
        style.setFillPattern(HSSFCellStyle.BIG_SPOTS);

//        HSSFCell cell = row.createCell((short) 1);
//        cell.setCellValue("X");
//        style = wb.createCellStyle();
        style.setFillForegroundColor(HSSFColor.ORANGE.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        return style;
    }

    //合并单元格
    private void mergeCell(XSSFSheet sheet) {
        CellRangeAddress region = new CellRangeAddress((short)1,(short)2,(short)3
                ,(short)4);//合并从第rowFrom行columnFrom列
        sheet.addMergedRegion(region);// 到rowTo行columnTo的区域
//得到所有区域
        sheet.getNumMergedRegions();
    }

    private void createExcel(XSSFWorkbook workbook, HttpServletResponse response) {
        try {
            XSSFCellStyle style = setStyle(workbook);

            XSSFSheet sheet = workbook.getSheetAt(0);
            // ------ 创建一行新的数据 ----------//
// 指定行索引，创建一行数据, 行索引为当前最后一行的行索引 + 1
            int currentLastRowIndex = sheet.getLastRowNum();
            int newRowIndex = currentLastRowIndex + 1;
            XSSFRow newRow = sheet.createRow(newRowIndex);
// 开始创建并设置该行每一单元格的信息，该行单元格的索引从 0 开始
            int cellIndex = 0;
// 创建一个单元格，设置其内的数据格式为字符串，并填充内容，其余单元格类同
            XSSFCell newNameCell = newRow.createCell(cellIndex++, Cell.CELL_TYPE_STRING);
            newNameCell.setCellValue("钱七");
            newNameCell.setCellStyle(style);
            XSSFCell newGenderCell = newRow.createCell(cellIndex++, Cell.CELL_TYPE_STRING);
            newGenderCell.setCellValue("女");
            XSSFCell newAgeCell = newRow.createCell(cellIndex++, Cell.CELL_TYPE_NUMERIC);
            newAgeCell.setCellValue(50);
            XSSFCell newWeightCell = newRow.createCell(cellIndex++, Cell.CELL_TYPE_NUMERIC);
            newWeightCell.setCellValue(68);
            XSSFCell newSalaryCell = newRow.createCell(cellIndex++, Cell.CELL_TYPE_NUMERIC);
            newSalaryCell.setCellValue(6000);

            //工作表锁定
            sheet.protectSheet("abcdefg");

//            EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
//            Encryptor enc = info.getEncryptor();
//            enc.confirmPassword("123456");
//            OPCPackage opc = OPCPackage.open(new File("abc"), PackageAccess.READ_WRITE);

            // 将最新的 Excel 数据写回到原始 Excel 文件（就是D盘那个 Excel 文件）中
// 首先要创建一个原始Excel文件的输出流对象！
            OutputStream out = response.getOutputStream();
            // 下载格式设置
            response.setContentType("APPLICATION/OCTET-STREAM");
            response.setHeader("Content-Disposition", "attachment; filename=\"employees1.xlsx\"");
//            FileOutputStream excelFileOutPutStream = new FileOutputStream("D:/employees1.xlsx");
// 将最新的 Excel 文件写入到文件输出流中，更新文件信息！
            workbook.write(out);
//            opc.save(out);
//            opc.close();
            // 执行 flush 操作， 将缓存区内的信息更新到文件上
            out.flush();
// 使用后，及时关闭这个输出流对象， 好习惯，再强调一遍！
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void createXlsExcel(HSSFWorkbook workbook, HttpServletResponse response, MultipartFile file) {
        try {
//            HSSFCellStyle style = setStyle(workbook);

            HSSFSheet sheet = workbook.getSheetAt(0);
            // ------ 创建一行新的数据 ----------//
// 指定行索引，创建一行数据, 行索引为当前最后一行的行索引 + 1
            int currentLastRowIndex = sheet.getLastRowNum();
            int newRowIndex = currentLastRowIndex + 1;
            HSSFRow newRow = sheet.createRow(newRowIndex);
// 开始创建并设置该行每一单元格的信息，该行单元格的索引从 0 开始
            int cellIndex = 0;
// 创建一个单元格，设置其内的数据格式为字符串，并填充内容，其余单元格类同
            HSSFCell newNameCell = newRow.createCell(cellIndex++, Cell.CELL_TYPE_STRING);
            newNameCell.setCellValue("钱七");
//            newNameCell.setCellStyle(style);
            HSSFCell newGenderCell = newRow.createCell(cellIndex++, Cell.CELL_TYPE_STRING);
            newGenderCell.setCellValue("女");
            HSSFCell newAgeCell = newRow.createCell(cellIndex++, Cell.CELL_TYPE_NUMERIC);
            newAgeCell.setCellValue(50);
            HSSFCell newWeightCell = newRow.createCell(cellIndex++, Cell.CELL_TYPE_NUMERIC);
            newWeightCell.setCellValue(68);
            HSSFCell newSalaryCell = newRow.createCell(cellIndex++, Cell.CELL_TYPE_NUMERIC);
            newSalaryCell.setCellValue(6000);

            //工作表锁定
            sheet.protectSheet("abcdefg");

//            EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
//            Encryptor enc = info.getEncryptor();
//            enc.confirmPassword("123456");
//            OPCPackage opc = OPCPackage.open(file.getInputStream());

            // 将最新的 Excel 数据写回到原始 Excel 文件（就是D盘那个 Excel 文件）中
// 首先要创建一个原始Excel文件的输出流对象！
            OutputStream out = response.getOutputStream();
            // 下载格式设置
            response.setContentType("APPLICATION/OCTET-STREAM");
            response.setHeader("Content-Disposition", "attachment; filename=\"employees1.xls\"");
//            FileOutputStream excelFileOutPutStream = new FileOutputStream("D:/employees1.xlsx");
// 将最新的 Excel 文件写入到文件输出流中，更新文件信息！
            workbook.write(out);
//            opc.save(out);
//            opc.close();
            // 执行 flush 操作， 将缓存区内的信息更新到文件上
            out.flush();
// 使用后，及时关闭这个输出流对象， 好习惯，再强调一遍！
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
