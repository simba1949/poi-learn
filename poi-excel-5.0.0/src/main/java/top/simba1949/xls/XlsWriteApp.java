package top.simba1949.xls;

import org.apache.poi.hssf.usermodel.*;
import top.simba1949.dao.UserDao;
import top.simba1949.domain.User;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Objects;

/**
 * @author anthony
 * @date 2021/3/14 18:56
 */
public class XlsWriteApp {

    public static void main(String[] args) throws IOException {
        String path = Objects.requireNonNull(XlsWriteApp.class.getClassLoader().getResource("")).getPath();
        String filePath = path + "user.xls";
        List<String> header = UserDao.getHeader();
        List<User> data = UserDao.getData();
        XlsWriteApp.writeCore(header, data, filePath);
    }

    public static void writeCore(List<String> headerList, List<User> userList, String xlsFilePath) throws IOException {
        // POI操作XLS文件对象
        HSSFWorkbook workbook = new HSSFWorkbook();
        // Sheet，工作簿对象
        HSSFSheet sheet = workbook.createSheet("xlsSheetName");

        for (int i = 0; i <= userList.size(); i++) {
            // Row，excel行对象
            HSSFRow row = sheet.createRow(i);
            for (int j = 0; j < headerList.size(); j++) {
                HSSFCell cell = row.createCell(j);
                if (0 == i){
                    // XLS文件头信息设置
                    // 设置单元格属性值
                    cell.setCellValue(headerList.get(j));
                    continue;
                }

                // XLS 数据设置
                User user = userList.get(i - 1);
                switch (j){
                    case 0: cell.setCellValue(user.getId()); break;
                    case 1: cell.setCellValue(user.getUsername()); break;
                    case 2: cell.setCellValue(user.getHeight().toPlainString()); break;
                    default:{
                        LocalDateTime birthday = user.getBirthday();
                        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
                        String format = birthday.format(dateTimeFormatter);
                        cell.setCellValue(format);
                    }
                }
            }

            // 文件输出流
            FileOutputStream fos = new FileOutputStream(xlsFilePath);
            // workbook 写入 Excel
            workbook.write(fos);
            // 关闭流
            fos.close();
            workbook.close();
        }
    }
}
