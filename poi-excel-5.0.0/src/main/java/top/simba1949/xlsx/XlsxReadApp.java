package top.simba1949.xlsx;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import top.simba1949.domain.User;
import top.simba1949.xls.XlsWriteApp;

import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * @author anthony
 * @date 2021/3/14 21:21
 */
public class XlsxReadApp {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        String path = Objects.requireNonNull(XlsWriteApp.class.getClassLoader().getResource("")).getPath();
        String filePath = path + "user.xlsx";

        List<User> users = readCore(filePath);

        users.forEach(System.out::println);
    }

    public static List<User> readCore(String xlsxPath) throws IOException, InvalidFormatException {
        List<User> result = new ArrayList<>();
        //
        OPCPackage opcPackage = OPCPackage.open(xlsxPath);
        // // POI操作XLSX文件对象
        XSSFWorkbook workbook = new XSSFWorkbook(opcPackage);
        // Sheet 对象
        XSSFSheet xssfSheet = workbook.getSheetAt(0);
        // 获取最后一个行的行号，如果行中间有空行也算空行(从0开始，空行为null)
        int lastRowNum = xssfSheet.getLastRowNum();
        int rowNumSum = lastRowNum + 1;

        // 获取row对象
        XSSFRow xssfRow = xssfSheet.getRow(0);
        // 获取最后一列的列号，如果列中间有空列也算空列（从1开始，最后一列列号也是总列数）
        short lastCellNum = xssfRow.getLastCellNum();
        int cellNumSum = lastCellNum;

        for (int i = 0; i < rowNumSum; i++) {
            XSSFRow dataRow = xssfSheet.getRow(i);
            if (0 == i || Objects.isNull(dataRow)){
                continue;
            }
            User user = new User();
            result.add(user);
            for (int j = 0; j < cellNumSum; j++) {
                XSSFCell cell = dataRow.getCell(j);
                switch (j){
                    case 0: {
                        double id = cell.getNumericCellValue();
                        user.setId((long)id);
                        break;
                    }
                    case 1:{
                        String username = cell.getStringCellValue();
                        user.setUsername(username);
                        break;
                    }
                    case 2:{
                        String heightStr = cell.getStringCellValue();
                        BigDecimal bigDecimal = new BigDecimal(heightStr);
                        user.setHeight(bigDecimal);
                        break;
                    }
                    default:{
                        String birthdayStr = cell.getStringCellValue();
                        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
                        LocalDateTime parse = LocalDateTime.parse(birthdayStr, dateTimeFormatter);
                        user.setBirthday(parse);
                    }
                }
            }
        }

        // 关闭所有流资源
        opcPackage.close();

        return result;
    }
}
