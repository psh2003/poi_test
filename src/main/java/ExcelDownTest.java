import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;

public class ExcelDownTest {
    public static void main(String[] args) throws Exception {
        // .xls 확장자 지원

        // HSSFWorkbook hssWb = null;
        // HSSFSheet hssSheet = null;
        // Row hssRow = null;
        // Cell hssCell = null;

        //.xlsx 확장자 지원

        XSSFWorkbook xssfWb = null;
        XSSFSheet xssfSheet = null;
        XSSFRow xssfRow = null;
        XSSFCell xssfCell = null;

        try {
            int rowNo = 0; // 행의 갯수
            // row(행) 순서 변수, cell(셀) 순서 변수
            int rowCount = 0;
            int cellCount = 0;

            xssfWb = new XSSFWorkbook(); //XSSFWorkbook 객체 생성
            xssfSheet = xssfWb.createSheet("거래처 정보"); // 워크시트 이름 설정

            // 폰트 스타일
            XSSFFont font = xssfWb.createFont();
            font.setFontName(HSSFFont.FONT_ARIAL); // 폰트 스타일
            font.setFontHeightInPoints((short) 20); // 폰트 크기

            //테이블 셀 스타일
            CellStyle cellStyle = xssfWb.createCellStyle();

            cellStyle.setAlignment(HorizontalAlignment.CENTER); // 정렬
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            // 테두리 선(좌, 우, 위, 아래)
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THIN);


            xssfRow = xssfSheet.createRow(rowNo); // 행 객체 추가

            xssfCell = xssfRow.createCell((short) 0); // 추가한 행에 셀 객체 추가
            xssfCell.setCellValue("NO"); // 데이터 입력

            xssfCell = xssfRow.createCell((short) 1); // 추가한 행에 셀 객체 추가
            xssfCell.setCellValue("수임 계약 정보"); // 데이터 입력

            xssfCell = xssfRow.createCell((short) 17); // 추가한 행에 셀 객체 추가
            xssfCell.setCellValue("거래처 정보"); // 데이터 입력

            xssfCell = xssfRow.createCell((short) 37); // 추가한 행에 셀 객체 추가
            xssfCell.setCellValue("본지점 정보"); // 데이터 입력

            xssfCell = xssfRow.createCell((short) 1); // 추가한 행에 셀 객체 추가

            xssfRow = xssfSheet.createRow(1); // 행 객체 추가

            xssfCell = xssfRow.createCell((short) 1); // 추가한 행에 셀 객체 추가
            xssfCell.setCellValue("수임구분"); // 데이터 입력


            for (int i = 6; i < 43; i++) {
                xssfCell = xssfRow.createCell((short) i); // 추가한 행에 셀 객체 추가
                xssfCell.setCellValue(i); // 데이터 입력

            }
            for (int i = 6; i < 43; i++) //autuSizeColumn after setColumnWidth setting!!
            {
                xssfSheet.autoSizeColumn(i);
                xssfSheet.setColumnWidth(i, (xssfSheet.getColumnWidth(i)) + 3000);
            }
            for (int i = 6; i < 43; i++) {
                xssfSheet.addMergedRegion(new CellRangeAddress(1, 2, i, i));
            }

            xssfRow = xssfSheet.createRow(2); // 행 객체 추가

            xssfCell = xssfRow.createCell((short) 1); // 추가한 행에 셀 객체 추가
            xssfCell.setCellValue("1"); // 데이터 입력

            xssfCell = xssfRow.createCell((short) 2); // 추가한 행에 셀 객체 추가
            xssfCell.setCellValue("2"); // 데이터 입력


            xssfCell = xssfRow.createCell((short) 3); // 추가한 행에 셀 객체 추가
            xssfCell.setCellValue("3"); // 데이터 입력


            xssfCell = xssfRow.createCell((short) 4); // 추가한 행에 셀 객체 추가
            xssfCell.setCellValue("4"); // 데이터 입력

            xssfCell = xssfRow.createCell((short) 5); // 추가한 행에 셀 객체 추가
            xssfCell.setCellValue("5"); // 데이터 입력


            xssfRow = xssfSheet.createRow(3); // 행 객체 추가
            for(int i=0;i<6;i++){

                xssfCell = xssfRow.createCell((short) i); // 추가한 행에 셀 객체 추가

                xssfSheet.addMergedRegion(new CellRangeAddress(3, 6, i, i));
                xssfCell.setCellValue(i); // 데이터 입력
            }

            for (int i = 6; i < 43; i++) {
                xssfCell = xssfRow.createCell((short) i); // 추가한 행에 셀 객체 추가
                xssfCell.setCellValue("test" + i); // 데이터 입력

            }
            for (int i = 6; i < 37; i++) {
                xssfSheet.addMergedRegion(new CellRangeAddress(3, 6, i, i));
            }

            xssfRow = xssfSheet.createRow(4); // 행 객체 추가.3

            for(int i=37;i<43;i++){
                xssfCell = xssfRow.createCell((short) i); // 추가한 행에 셀 객체 추가
                xssfCell.setCellValue(i-36); // 데이터 입력

            }

            xssfRow = xssfSheet.createRow(5); // 행 객체 추가

            for (int i = 36; i < 43; i++) {
                xssfCell = xssfRow.createCell((short) i); // 추가한 행에 셀 객체 추가


            }
            xssfRow = xssfSheet.createRow(6); // 행 객체 추가

            for (int i = 36; i < 43; i++) {
                xssfCell = xssfRow.createCell((short) i); // 추가한 행에 셀 객체 추가


            }

            for(int i=0;i<7;i++){
                for(int j=0;j<43;j++){
                    xssfRow = xssfSheet.getRow(i);
                    xssfCell = xssfRow.getCell(j);
                    if(xssfCell==null) xssfCell = xssfRow.createCell((short) j); // 추가한 행에 셀 객체 추가
                    if(i==0&&j==1) xssfCell.setCellValue("수임 계약 정보"); // 데이터 입력
                    xssfCell.setCellStyle(cellStyle);
                }
            }
            xssfSheet.addMergedRegion(new CellRangeAddress(0, 2, 0, 0)); //첫행, 마지막행, 첫열, 마지막열 병합
            xssfSheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 16));
            xssfSheet.addMergedRegion(new CellRangeAddress(0, 0, 17, 36));
            xssfSheet.addMergedRegion(new CellRangeAddress(0, 0, 37, 42));
            xssfSheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 5));
            String localFile = "/Users/Neo/Desktop/" + "excelDownTest" + ".xlsx";

            File file = new File(localFile);
            FileOutputStream fos = null;
            fos = new FileOutputStream(file);
            xssfWb.write(fos);

            if (fos != null) fos.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
