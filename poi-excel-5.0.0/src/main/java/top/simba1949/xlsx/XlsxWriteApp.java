package top.simba1949.xlsx;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;
import top.simba1949.dao.UserDao;
import top.simba1949.domain.User;
import top.simba1949.xls.XlsWriteApp;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Objects;

/**
 * @author anthony
 * @date 2021/3/14 20:55
 */
public class XlsxWriteApp {

    public static void main(String[] args) throws IOException {
        String path = Objects.requireNonNull(XlsWriteApp.class.getClassLoader().getResource("")).getPath();
        String filePath = path + "user.xlsx";
        List<String> header = UserDao.getHeader();
        List<User> data = UserDao.getData();
        XlsxWriteApp.writeCore(header, data, filePath);
    }

    public static void writeCore(List<String> headerList, List<User> userList, String xlsxFilePath) throws IOException {
        // POI操作XLSX文件对象
        XSSFWorkbook workbook = new XSSFWorkbook();
        // Sheet 对象
        XSSFSheet xssfSheet = workbook.createSheet("xlsxSheetName");

        // 创建单元格格式
        XSSFCellStyle cellStyle4Header = workbook.createCellStyle();
        cellStyle4Header.setFont(getFont4Header(workbook));

        for (int i = 0; i <= userList.size(); i++) {
            // XSSFRow 行对象
            XSSFRow xssfRow = xssfSheet.createRow(i);
            for (int j = 0; j < headerList.size(); j++) {
                // XSSFCell 单元格对象
                XSSFCell xssfCell = xssfRow.createCell(j);
                if (0 == i){
                    // 首行信息设置
                    // 设置单元格属性值
                    xssfCell.setCellValue(headerList.get(j));
                    xssfCell.setCellStyle(cellStyle4Header);
                    continue;
                }

                // XLSX 数据设置
                User user = userList.get(i - 1);
                switch (j){
                    case 0: xssfCell.setCellValue(user.getId()); break;
                    case 1: xssfCell.setCellValue(user.getUsername()); break;
                    case 2: xssfCell.setCellValue(user.getHeight().toPlainString()); break;
                    default:{
                        LocalDateTime birthday = user.getBirthday();
                        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
                        String format = birthday.format(dateTimeFormatter);
                        xssfCell.setCellValue(format);
                    }
                }
            }

            // 文件输出流
            FileOutputStream fos = new FileOutputStream(xlsxFilePath);
            // workbook 写入 Excel
            workbook.write(fos);
            // TODO 关闭会导致无法写入数据
            // fos.close();
            // workbook.close();
        }
    }

    /**
     * 获取字体
     * @param workbook
     * @return
     */
    public static XSSFFont getFont4Header(XSSFWorkbook workbook){
        //  创建字体
        XSSFFont font4Header = workbook.createFont();
        // 设置字体颜色
        font4Header.setColor(IndexedColors.RED.getIndex());
        // 设置字体样式，支持中文字体名称
        font4Header.setFontName("楷体");
        // 设置字体是否加粗
        font4Header.setBold(true);
        // 设置字体大小
        font4Header.setFontHeight(50);

        return font4Header;
    }
}
