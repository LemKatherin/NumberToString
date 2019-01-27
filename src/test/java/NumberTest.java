import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;


import org.testng.Assert;
import org.testng.ITestResult;
import org.testng.annotations.*;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;


public class NumberTest {

    String num = "";
    String numStr = "";
    static int numtest = 0;


    // тест для проверки одного значения
    @Test
    public void getString() throws Exception {

        String testNumber = "123";
        String testNumberText = "сто двадцать три";

        Number expected = new Number(testNumber);
        expected.toString();
        Assert.assertEquals(testNumberText, expected.getString(), "Ошибка в переводе числа " + testNumber);

    }

    @DataProvider
    public Object[][] testData() throws IOException {
        ArrayList<String[]> dataList = new ArrayList<>();
        InputStream in = null;
        HSSFWorkbook wb = null;

        try {
            in = new FileInputStream("TestData.xls");
            wb = new HSSFWorkbook(in);
        } catch (IOException e) {
            e.printStackTrace();
        }
        Sheet sheet = wb.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        int i = 0;
        while (it.hasNext()){
            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            num = "";
            numStr = "";
            int k = 0;
            while (cells.hasNext() && k < 2) {

                Cell cell = cells.next();
                CellType cellType = cell.getCellType();

                switch (cellType) {
                    case STRING:
                        numStr += cell.getStringCellValue();
                        break;
                    case NUMERIC:
                        num += cell.getNumericCellValue();
                        num = num.substring(0,num.indexOf('.'));
                        break;
                    default:
                        break;
                }
                k++;
            }
            dataList.add(new String[]{num, numStr});
            i++;
        }

        Object[][] data = new Object[i][2];
        int k = 0;
        for(String[] j : dataList)
        {
            data[k] = j;
            k++;
        }
        in.close();
        return data;
    }

    @Test (dataProvider = "testData")
    public void fromExcel(String n, String s) throws Exception{
        Number expected;
        expected = new Number(n);
        expected.toString();
        numtest++;
        Assert.assertEquals(s, expected.getString(), "Ошибка в переводе числа " + n);

    }

            // тесты для проврки правильности вводимого числа

    @Test (expectedExceptions = Exception.class)
    public void emptyString() throws Exception{
        Number N = new Number("");
    }

    @Test (expectedExceptions = Exception.class)
    public void symbolsInString() throws Exception{
        Number N = new Number("123-4");
    }

    @Test (expectedExceptions = Exception.class)
    public void tooHigh() throws Exception{
        Number N = new Number("1111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111");
    }

    @Test (expectedExceptions = Exception.class)
    public void startsWithZero() throws Exception{
        Number N = new Number("0123");
    }

    @AfterMethod
    public void testSuccess(ITestResult result) throws IOException {
        if(result.getName() == "fromExcel"){
            OutputStream out = null;
            InputStream in = null;
            HSSFWorkbook wb = null;
            try {
                in = new FileInputStream("TestData.xls");
                wb = new HSSFWorkbook(in);

            } catch (IOException e) {
                e.printStackTrace();
            }
            Sheet sheet = wb.getSheetAt(0);
            if(result.isSuccess()) {
                if(sheet.getRow(numtest - 1).getCell(2) == null)
                    sheet.getRow(numtest - 1).createCell(2);
                sheet.getRow(numtest - 1).getCell(2).setCellValue(true);
            }
            else {
                if(sheet.getRow(numtest - 1).getCell(2) == null)
                    sheet.getRow(numtest - 1).createCell(2);
                sheet.getRow(numtest - 1).getCell(2).setCellValue(false);
            }
            in.close();


            out = new FileOutputStream("TestData.xls");
            wb.write(out);
            out.close();
        }
        else if((result.getName() == "startsWithZero" || result.getName() == "tooHigh"  ||
                result.getName() == "symbolsInString"  || result.getName() == "emptyString") &&
                result.isSuccess()){
            System.out.println("Произошла ошибка при выполнении негативного теста " + result.getName());
        }

    }

}