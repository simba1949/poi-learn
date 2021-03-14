package top.simba1949.xls;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import top.simba1949.domain.User;

import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * @author anthony
 * @date 2021/3/14 20:18
 */
public class XlsReadApp {

    public static void main(String[] args) throws IOException {
        String path = Objects.requireNonNull(XlsWriteApp.class.getClassLoader().getResource("")).getPath();
        String filePath = path + "user.xls";

        List<User> users = readCore(filePath);

        users.forEach(System.out::println);
    }

    public static List<User> readCore(String xlsPath) throws IOException {
        List<User> result = new ArrayList<>();
        // XSSFWorkbook 对象
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(xlsPath));
        // sheet，工作簿对象
        HSSFSheet sheet = workbook.getSheetAt(0);
        // 获取最后一个行的行号，如果行中间有空行也算空行(从0开始，空行为null)
        int lastRowNum = sheet.getLastRowNum();
        int rowNumSum = lastRowNum + 1;

        // 获取row对象
        HSSFRow row = sheet.getRow(0);
        // 获取最后一列的列号，如果列中间有空列也算空列（从1开始，最后一列列号也是总列数）
        short lastCellNum = row.getLastCellNum();
        int cellNumSum = lastCellNum;
        for (int i = 0; i < rowNumSum; i++) {
            HSSFRow dataRow = sheet.getRow(i);
            if (0 == i || Objects.isNull(dataRow)){
                // 第一行是标题需要跳过，空行需要跳过
                continue;
            }
            User user = new User();
            result.add(user);
            for (int j = 0; j < cellNumSum; j++) {
                HSSFCell cell = dataRow.getCell(j);
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

        return result;
    }
}
