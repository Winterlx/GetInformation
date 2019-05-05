import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class Main {

    List<Data> data = new ArrayList<>();

    public static void main(String[] args) throws Exception {

        Main main = new Main();

    }
    // TODO: 2019/4/23 根据时间生成文件名


    public Main() {

        String[] title = new String[]{"Tag ID","KeyStore","Value","Password","Invalid","Time Stamp"};

        try {

            File file = new File("D:\\NewCashProject\\storage.txt");
            List<String> stringList = FileUtils.readLines(file, "UTF-8");
            for (int i = 0; i < stringList.size(); i++) {
                System.out.println("data : " + i + " " + stringList.get(i));

                String[] str = stringList.get(i).split("/");


                Data dataTemp = new Data();
                System.out.println(str[0]);
                System.out.println(str[1]);
                System.out.println(str[2]);
                System.out.println(str[3]);
                System.out.println(str[4]);

                dataTemp.setTimeStamp(str[0]);
                dataTemp.setPassword(str[1]);
                dataTemp.setKeyStore(str[2]);
                dataTemp.setValue(str[3]);
                dataTemp.setTagID(str[4]);

                data.add(dataTemp);
            }

            writeEmployeeListToExcel("D:\\NewCashProject\\Data.xls",title,data,"NO.1");



        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    /**
     * 将List集合数据写入excel（单个sheet）
     *
     * @param filePath   文件路径
     * @param excelTitle 文件表头
     * @param data       要写入的数据集合
     * @param sheetName  sheet名称
     */
    public static void writeEmployeeListToExcel(String filePath, String[] excelTitle, List<Data> data, String sheetName) {
        System.out.println("开始写入文件>>>>>>>>>>>>");
        Workbook workbook = null;

        if (filePath.toLowerCase().endsWith("xls")) {//2003
            workbook = new HSSFWorkbook();
        }else {
            //			logger.debug("invalid file name,should be xls or xlsx");
        }

        /*if (filePath.toLowerCase().endsWith("xls")) {//2003
            workbook = new XSSFWorkbook();
        }else
        if (filePath.toLowerCase().endsWith("xlsx")) {//2007
            workbook = new HSSFWorkbook();
        } else {
//			logger.debug("invalid file name,should be xls or xlsx");
        }*/
        //create sheet
        Sheet sheet = workbook.createSheet(sheetName);
        int rowIndex = 0;//标识位，用于标识sheet的行号
        //遍历数据集，将其写入excel中
        try {
            //写表头数据
            Row titleRow = sheet.createRow(rowIndex);
            for (int i = 0; i < excelTitle.length; i++) {
                //创建表头单元格,填值
                titleRow.createCell(i).setCellValue(excelTitle[i]);
            }
            System.out.println("表头写入完成>>>>>>>>");
            rowIndex++;
            //循环写入主表数据
            for (Iterator<Data> employeeIter = data.iterator(); employeeIter.hasNext(); ) {
                Data employee = employeeIter.next();
                //create sheet row
                Row row = sheet.createRow(rowIndex);
                //create sheet coluum(单元格)
                Cell cell0 = row.createCell(0);
                cell0.setCellValue(employee.getTagID());
                Cell cell1 = row.createCell(1);
                cell1.setCellValue(employee.getValue());
                Cell cell2 = row.createCell(2);
                cell2.setCellValue(employee.getKeyStore());
                Cell cell3 = row.createCell(3);
                cell3.setCellValue(employee.getPassword());
                Cell cell4 = row.createCell(4);
                cell4.setCellValue(employee.getInvalid());
                Cell cell5 = row.createCell(5);
                cell5.setCellValue(employee.getTimeStamp());
                rowIndex++;
            }
            System.out.println("主表数据写入完成>>>>>>>>");
            FileOutputStream fos = new FileOutputStream(filePath);
            workbook.write(fos);
            fos.close();
            System.out.println(filePath + "写入文件成功>>>>>>>>>>>");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 读取Excel2003的主表数据 （单个sheet）
     *
     * @param filePath
     * @return
     */
    private static List<Data> readFromXLS2003(String filePath) {
        File excelFile = null;// Excel文件对象
        InputStream is = null;// 输入流对象
        String cellStr = null;// 单元格，最终按字符串处理
        List<Data> employeeList = new ArrayList<Data>();// 返回封装数据的List
        Data employee = null;// 每一个雇员信息对象
        try {
            excelFile = new File(filePath);
            is = new FileInputStream(excelFile);// 获取文件输入流
            HSSFWorkbook workbook2003 = new HSSFWorkbook(is);// 创建Excel2003文件对象
            HSSFSheet sheet = workbook2003.getSheetAt(0);// 取出第一个工作表，索引是0
            // 开始循环遍历行，表头不处理，从1开始
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                HSSFRow row = sheet.getRow(i);// 获取行对象
                employee = new Data();// 实例化Student对象
                if (row == null) {// 如果为空，不处理
                    continue;
                }
                // 循环遍历单元格
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    HSSFCell cell = row.getCell(j);// 获取单元格对象
                    if (cell == null) {// 单元格为空设置cellStr为空串
                        cellStr = "";
                    }/* else if (cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {// 对布尔值的处理
                        cellStr = String.valueOf(cell.getBooleanCellValue());
                    } else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {// 对数字值的处理
                        cellStr = cell.getNumericCellValue() + "";
                    }*/ else {// 其余按照字符串处理
                        cellStr = cell.getStringCellValue();
                    }
                    // 下面按照数据出现位置封装到bean中
                    if (j == 0) {
                        employee.setTagID(cellStr);
                    } else if (j == 1) {
                        employee.setValue(cellStr);
                    } else if (j == 2) {
                        employee.setKeyStore(cellStr);
                    } else if (j == 3) {
                        employee.setPassword(cellStr);
                    } else if (j == 4) {
                        employee.setInvalid(cellStr);
                    } else {
                        employee.setTimeStamp(cellStr);
                    }
                }
                employeeList.add(employee);// 数据装入List
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {// 关闭文件流
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return employeeList;
    }

    /**
     * 读取Excel2003的表头
     *
     * @param filePath 需要读取的文件路径
     * @return
     */
    public static String[] readHeaderFromXLS2003(String filePath) {
        String[] excelTitle = null;
        FileInputStream is = null;
        try {
            File excelFile = new File(filePath);
            is = new FileInputStream(excelFile);
            HSSFWorkbook workbook2003 = new HSSFWorkbook(is);
            //循环读取工作表
            for (int i = 0; i < workbook2003.getNumberOfSheets(); i++) {
                HSSFSheet hssfSheet = workbook2003.getSheetAt(i);
                //*************获取表头是start*************
                HSSFRow sheetRow = hssfSheet.getRow(i);
                excelTitle = new String[sheetRow.getLastCellNum()];
                for (int k = 0; k < sheetRow.getLastCellNum(); k++) {
                    HSSFCell hssfCell = sheetRow.getCell(k);
                    excelTitle[k] = hssfCell.getStringCellValue();
//		            	System.out.println(excelTitle[k] + " ");
                }
                //*************获取表头end*************
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {// 关闭文件流
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return excelTitle;
    }


}
