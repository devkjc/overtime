package com.fine.overtime.util;

import com.fine.overtime.domain.OverTimeGroup;
import com.fine.overtime.domain.OverTimePeople;
import com.fine.overtime.domain.OverTimeReceipt;
import com.fine.overtime.repo.PeopleRepo;
import com.fine.overtime.repo.ReceiptRepo;
import lombok.RequiredArgsConstructor;
import org.apache.poi.hssf.util.HSSFColor.GREY_25_PERCENT;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

@Service
@RequiredArgsConstructor
public class ExcelOverTime {

    private final ReceiptRepo receiptRepo;
    private final PeopleRepo peopleRepo;
    private final FileUtils fileUtils;

    public void execute(OverTimeGroup group, HttpSession session, HttpServletRequest request, HttpServletResponse response) throws Exception {

        String path = createExcel(group, session);
        String fileName = group.getGroup_name()+"_야근식대.zip";
        fileUtils.createZipFile(path,fileName);
        fileUtils.fileDown(path,fileName,request,response,session);

        fileUtils.directoryDelete(path);
    }

    private String createExcel(OverTimeGroup group, HttpSession session) {

        String defaultPath = session.getServletContext().getRealPath("/"); //서버 기본 경로(프로젝트X)
        String path = defaultPath + "upload" + File.separator + "temp" + File.separator;

        int X = 0;
        String job = null;
        String name = null;
        String role = null;

        ArrayList<OverTimeReceipt> receipts = receiptRepo.findByGroup(group);

        // 날짜별로 엑셀 파일 생성
        for (OverTimeReceipt dto : receipts) {

            String date = dto.getReceipt_date();
            String overtime_name = dto.getReceipt_name();
            int overtime_money = dto.getReceipt_amount();

            // 야근 식대별 인원 수 계산
            if (overtime_money <= 10000) {
                X = 1;
            } else {
                if (overtime_money % 10000 == 0) {
                    X = overtime_money / 10000;
                } else {
                    X = (overtime_money / 10000) + 1;
                }
            }

            //----------------------------------------------------------
            //WorkBook 생성 및 시트 , 변수 생성
            Workbook xlsxWb = new XSSFWorkbook(); // Excel 2007 이상
            Sheet sheet1 = xlsxWb.createSheet("야근일지");

            sheet1.setColumnWidth(1, 150 * 20);
            sheet1.setColumnWidth(2, 150 * 20);
            sheet1.setColumnWidth(3, 300 * 20);
            sheet1.setColumnWidth(4, 150 * 20);
            sheet1.setColumnWidth(5, 300 * 20);

            sheet1.setDefaultRowHeight((short) (30 * 20));

            //폰트 설정
            Font Black11 = xlsxWb.createFont();
            Black11.setFontName("나눔고딕"); //글씨체
            Black11.setFontHeight((short) (10.5 * 20)); //사이즈

            //폰트 설정
            Font Title = xlsxWb.createFont();
            Title.setFontName("나눔고딕"); //글씨체
            Title.setFontHeight((short) (14 * 20)); //사이즈
            Title.setBold(true); // 볼드

            Row row = null;
            Cell cell = null;

            //----------------------------------------------------------
            //스타일 설정 (가운데 정렬 , 모든테두리)
            CellStyle centerAlign = xlsxWb.createCellStyle();
            centerAlign.setAlignment(HorizontalAlignment.CENTER);
            centerAlign.setVerticalAlignment(VerticalAlignment.CENTER);

            centerAlign.setBorderRight(BorderStyle.THIN);
            centerAlign.setBorderLeft(BorderStyle.THIN);
            centerAlign.setBorderTop(BorderStyle.THIN);
            centerAlign.setBorderBottom(BorderStyle.THIN);

            centerAlign.setFont(Black11);

            //----------------------------------------------------------
            //스타일 설정 (가운데 정렬 , 모든테두리 , 볼드 , 14pt)
            CellStyle title = xlsxWb.createCellStyle();
            title.setAlignment(HorizontalAlignment.CENTER);
            title.setVerticalAlignment(VerticalAlignment.CENTER);

            title.setBorderRight(BorderStyle.THIN);
            title.setBorderLeft(BorderStyle.THIN);
            title.setBorderTop(BorderStyle.THIN);
            title.setBorderBottom(BorderStyle.THIN);

            title.setFont(Title);

            //----------------------------------------------------------
            //----------------------------------------------------------
            //스타일 설정 (가운데 정렬 , 모든테두리 , 고정컬럼(회색))
            CellStyle CenterGray = xlsxWb.createCellStyle();
            CenterGray.setAlignment(HorizontalAlignment.CENTER);
            CenterGray.setVerticalAlignment(VerticalAlignment.CENTER);

            CenterGray.setFillForegroundColor(GREY_25_PERCENT.index);
            CenterGray.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            CenterGray.setBorderRight(BorderStyle.THIN);
            CenterGray.setBorderLeft(BorderStyle.THIN);
            CenterGray.setBorderTop(BorderStyle.THIN);
            CenterGray.setBorderBottom(BorderStyle.THIN);

            CenterGray.setFont(Black11);

            //-------------------------------------------------------------------------------------------
            //맨 위 병합 셀 생성
            row = sheet1.createRow(1);

            // 첫 번째 줄에 Cell 설정하기-------------

            cell = row.createCell(1);
            cell.setCellValue("야 근 일 지");
            cell.setCellStyle(title);
            cell = row.createCell(2);
            cell.setCellStyle(title);
            cell = row.createCell(3);
            cell.setCellStyle(title);
            cell = row.createCell(4);
            cell.setCellStyle(title);
            cell = row.createCell(5);
            cell.setCellStyle(title);
            // 두 번째 줄에 Cell 설정하기-------------
            row = sheet1.createRow(2);
            cell = row.createCell(1);
            cell.setCellStyle(title);
            cell = row.createCell(2);
            cell.setCellStyle(title);
            cell = row.createCell(3);
            cell.setCellStyle(title);
            cell = row.createCell(4);
            cell.setCellStyle(title);
            cell = row.createCell(5);
            cell.setCellStyle(title);

            // 상단 제목 병합
            sheet1.addMergedRegion(new CellRangeAddress(1, 2, 1, 5)); //startRow,endRow,startCol,endCol
            //--------------------------------------------------------------------------------------------------
            //세번째 줄
            row = sheet1.createRow(3);

            cell = row.createCell(1);
            cell.setCellValue("연구책임자");
            cell.setCellStyle(CenterGray);

            cell = row.createCell(2);
            cell.setCellValue("소속");
            cell.setCellStyle(CenterGray);

            cell = row.createCell(3);
            cell.setCellValue(group.getGroup_officer_dept());
            cell.setCellStyle(centerAlign);

            cell = row.createCell(4);
            cell.setCellValue("성명");
            cell.setCellStyle(CenterGray);

            cell = row.createCell(5);
            cell.setCellValue(group.getGroup_officer_name());
            cell.setCellStyle(centerAlign);
            //--------------------------------------------------------------------------------------------------
            //네번째 줄

            row = sheet1.createRow(4);

            cell = row.createCell(1);
            cell.setCellValue("지원 기관");
            cell.setCellStyle(CenterGray);

            cell = row.createCell(2);
            cell.setCellValue(group.getGroup_support());
            cell.setCellStyle(centerAlign);
            cell = row.createCell(3);
            cell.setCellStyle(centerAlign);

            cell = row.createCell(4);
            cell.setCellValue("연구 기간");
            cell.setCellStyle(CenterGray);

            cell = row.createCell(5);
            cell.setCellValue(group.getGroup_period());
            cell.setCellStyle(centerAlign);

            sheet1.addMergedRegion(new CellRangeAddress(4, 4, 2, 3)); //startRow,endRow,startCol,endCol
            //--------------------------------------------------------------------------------------------------
            //다섯번째 줄

            row = sheet1.createRow(5);

            cell = row.createCell(1);
            cell.setCellValue("연구 과제명");
            cell.setCellStyle(CenterGray);

            cell = row.createCell(2);
            cell.setCellValue(group.getGroup_subject_name());
            cell.setCellStyle(centerAlign);
            cell = row.createCell(3);
            cell.setCellStyle(centerAlign);
            cell = row.createCell(4);
            cell.setCellStyle(centerAlign);
            cell = row.createCell(5);
            cell.setCellStyle(centerAlign);
            sheet1.addMergedRegion(new CellRangeAddress(5, 5, 2, 5)); //startRow,endRow,startCol,endCol

            //--------------------------------------------------------------------------------------------------
            //여섯번째 줄

            row = sheet1.createRow(6);

            cell = row.createCell(1);
            cell.setCellValue("야근 장소");
            cell.setCellStyle(CenterGray);

            cell = row.createCell(2);
            cell.setCellValue(group.getGroup_place());
            cell.setCellStyle(centerAlign);
            cell = row.createCell(3);
            cell.setCellStyle(centerAlign);

            cell = row.createCell(4);
            cell.setCellValue("야근일");
            cell.setCellStyle(CenterGray);

            cell = row.createCell(5);
            cell.setCellValue(date);
            cell.setCellStyle(centerAlign);
            sheet1.addMergedRegion(new CellRangeAddress(6, 6, 2, 3)); //startRow,endRow,startCol,endCol

            //--------------------------------------------------------------------------------------------------
            //--------------------------------------------------------------------------------------------------
            //일곱번째 줄

            row = sheet1.createRow(7);

            cell = row.createCell(1);
            cell.setCellValue("야근자 명단");
            cell.setCellStyle(CenterGray);

            cell = row.createCell(2);
            cell.setCellValue("직급");
            cell.setCellStyle(CenterGray);

            cell = row.createCell(3);
            cell.setCellValue("성명");
            cell.setCellStyle(CenterGray);

            cell = row.createCell(4);
            cell.setCellValue("근무내용");
            cell.setCellStyle(CenterGray);
            cell = row.createCell(5);
            cell.setCellStyle(CenterGray);

            sheet1.addMergedRegion(new CellRangeAddress(7, 7, 4, 5)); //startRow,endRow,startCol,endCol

            //--------------------------------------------------------------------------------------------------
            //야근자 명단.

            ArrayList<OverTimePeople> overtime_person = peopleRepo.findByGroup(group);
            int aa = 8;

            for (OverTimePeople person : overtime_person) {

                job = person.getPeople_position();
                name = person.getPeople_name();
                role = person.getPeople_duties();

                System.out.println(job + " " + name + " " + role);

                row = sheet1.createRow(aa);

                cell = row.createCell(1);
                cell.setCellValue("야근자 명단");
                cell.setCellStyle(CenterGray);

                cell = row.createCell(2);
                cell.setCellValue(job);
                cell.setCellStyle(centerAlign);

                cell = row.createCell(3);
                cell.setCellValue(name);
                cell.setCellStyle(centerAlign);

                cell = row.createCell(4);
                cell.setCellValue(role);
                cell.setCellStyle(centerAlign);
                cell = row.createCell(5);
                cell.setCellStyle(centerAlign);

                // 근무내용 가로 병합
                sheet1.addMergedRegion(new CellRangeAddress(aa, aa, 4, 5)); //startRow,endRow,startCol,endCol
                aa++;
            }

            for (int i = 0; i < 10 - overtime_person.size(); i++) {

                row = sheet1.createRow(aa);

                cell = row.createCell(1);
                cell.setCellValue("야근자 명단");
                cell.setCellStyle(CenterGray);

                cell = row.createCell(2);
                cell.setCellStyle(centerAlign);

                cell = row.createCell(3);
                cell.setCellStyle(centerAlign);

                cell = row.createCell(4);
                cell.setCellStyle(centerAlign);
                cell = row.createCell(5);
                cell.setCellStyle(centerAlign);

                // 근무내용 가로 병합
                sheet1.addMergedRegion(new CellRangeAddress(aa, aa, 4, 5)); //startRow,endRow,startCol,endCol
                aa++;
            }

            //--------------------------------------------------------------------------------------------------
            // 야근자 명단 세로 병합
            sheet1.addMergedRegion(new CellRangeAddress(8, 17, 1, 1)); //startRow,endRow,startCol,endCol

            File dirF = new File(path);

            // excel 파일 저장
            try {
                if (!dirF.exists()) {
                    dirF.mkdirs();
                }
                File xlsFile = new File(path + File.separator + date + "_" + overtime_name + "_" + overtime_money + "_야근일지.xlsx");
                FileOutputStream fileOut = new FileOutputStream(xlsFile);
                xlsxWb.write(fileOut);

            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                try {
                    xlsxWb.close();
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }

        }

        return path;
    }


}