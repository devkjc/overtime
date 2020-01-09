package com.fine.overtime.util;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Locale;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import com.fine.overtime.domain.OverTimeGroup;
import com.fine.overtime.domain.OverTimePeople;
import com.fine.overtime.domain.OverTimeReceipt;
import com.fine.overtime.repo.GroupRepo;
import com.fine.overtime.repo.PeopleRepo;
import com.fine.overtime.repo.ReceiptRepo;
import org.apache.poi.hssf.util.HSSFColor.GREY_25_PERCENT;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.util.FileCopyUtils;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

@Service
public class ExcelOverTime {

    private ReceiptRepo receiptRepo;
    private PeopleRepo peopleRepo;
    private GroupRepo groupRepo;

    public ExcelOverTime(ReceiptRepo receiptRepo, PeopleRepo peopleRepo, GroupRepo groupRepo) {
        this.receiptRepo = receiptRepo;
        this.peopleRepo = peopleRepo;
        this.groupRepo = groupRepo;
    }

    public void execute(OverTimeGroup group, HttpSession session, HttpServletRequest request, HttpServletResponse response) throws Exception {

        String path = createExcel(group, session);
        String fileName = group.getGroup_name()+"_야근식대.zip";
        createZipFile(path,fileName);
        fileDown(path,fileName,request,response,session);

        directoryDelete(path);
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

    /**
     * 디렉토리 및 파일을 압축한다.
     *
     * @param path     압축할 디렉토리 및 파일
     * @param fileName 압축파일의 이름
     */

    private void createZipFile(String path, String fileName) {

        File dir = new File(path);
        String[] list = dir.list();
        String _path;

        if (!dir.canRead() || !dir.canWrite()) {
            return;
        }

        int len = list.length;

        if (path.charAt(path.length() - 1) != '/') {
            _path = path + "/";
        } else {
            _path = path;
        }

        try {
            ZipOutputStream zip_out = new ZipOutputStream(new BufferedOutputStream(new FileOutputStream(path + "/" + fileName), 2048));

            for (int i = 0; i < len; i++) {
                zip_folder("", new File(_path + list[i]), zip_out, path);
            }

            zip_out.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {


        }
    }

    /**
     * ZipOutputStream를 넘겨 받아서 하나의 압축파일로 만든다.
     *
     * @param parent 상위폴더명
     * @param file   압축할 파일
     * @param zout   압축전체스트림
     * @throws IOException
     */

    private void zip_folder(String parent, File file, ZipOutputStream zout, String toPath) throws IOException {
        byte[] data = new byte[2048];
        int read;

        if (file.isFile()) {
            ZipEntry entry = new ZipEntry(parent + file.getName());
            zout.putNextEntry(entry);
            BufferedInputStream instream = new BufferedInputStream(new FileInputStream(file));

            while ((read = instream.read(data, 0, 2048)) != -1)
                zout.write(data, 0, read);

            zout.flush();
            zout.closeEntry();
            instream.close();

        } else if (file.isDirectory()) {
            String parentString = file.getPath().replace(toPath, "");
            parentString = parentString.substring(0, parentString.length() - file.getName().length());
            ZipEntry entry = new ZipEntry(parentString + file.getName() + "/");
            zout.putNextEntry(entry);

            String[] list = file.list();
            if (list != null) {
                int len = list.length;
                for (int i = 0; i < len; i++) {
                    zip_folder(entry.getName(), new File(file.getPath() + "/" + list[i]), zout, toPath);
                }
            }
        }
    }

    private void directoryDelete(String path) {
        File folder = new File(path);
        try {
            while (folder.exists()) {
                File[] folder_list = folder.listFiles(); //파일리스트 얻어오기

                for (int j = 0; j < folder_list.length; j++) {
                    folder_list[j].delete(); //파일 삭제
                }
                if (folder_list.length == 0 && folder.isDirectory()) {
                    folder.delete(); //대상폴더 삭제
                }
            }
        } catch (Exception e) {
            e.getStackTrace();
        }
    }

    private void fileDown(String fileUrl, String oriFileName, HttpServletRequest request, HttpServletResponse response, HttpSession session) throws Exception {

        request.setCharacterEncoding("UTF-8");
        System.out.println("파일이름: " + oriFileName);

        //파일 업로드된 경 로
        try {

            fileUrl = fileUrl + oriFileName;

            System.out.println("fileUrl:" + fileUrl);

            //실제 내보낼 파일명
            InputStream in = null;
            OutputStream os = null;
            File file = null;
            boolean skip = false;
            String client = "";

            //파일을 읽어 스트림에 담기
            try {
                file = new File(fileUrl);
                in = new FileInputStream(file);
            } catch (FileNotFoundException fe) {
                skip = true;
            }

            client = request.getHeader("User-Agent");

            //파일 다운로드 헤더 지정
            response.reset();
            response.setContentType("application/octet-stream");
            response.setHeader("Content-Description", "JSP Generated Data");

            if (!skip) {
                String[] invalidName = {"\\\\", "/", ":", "[*]", "[?]", "\"", "<", ">", "[|]", "&", " "}; // 윈도우 파일명으로 사용할수 없는 문자

                for (int i = 0; i < invalidName.length; i++) {
                    oriFileName = oriFileName.replaceAll(invalidName[i], "_");
                }

                System.out.println("변환된 파일이름: " + oriFileName);

                // IE
                if (client.indexOf("MSIE") != -1) {
                    response.setHeader("Content-Disposition", "attachment; filename=\""
                            + java.net.URLEncoder.encode(oriFileName, "UTF-8").replaceAll("\\+", "\\ ") + "\"");
                    // IE 11 이상.
                } else if (client.indexOf("Trident") != -1) {
                    response.setHeader("Content-Disposition", "attachment; filename=\""
                            + java.net.URLEncoder.encode(oriFileName, "UTF-8").replaceAll("\\+", "\\ ") + "\"");
                } else {
                    // 한글 파일명 처리
                    response.setHeader("Content-Disposition",
                            "attachment; filename=\"" + new String(oriFileName.getBytes("UTF-8"), "ISO8859_1") + "\"");
                    response.setHeader("Content-Type", "application/octet-stream; charset=utf-8");
                }
                response.setHeader("Content-Length", "" + file.length());
                os = response.getOutputStream();
                byte b[] = new byte[(int) file.length()];
                int leng = 0;
                while ((leng = in.read(b)) > 0) {
                    os.write(b, 0, leng);
                }

            } else {
                response.setContentType("text/html;charset=UTF-8");

                PrintWriter writer = response.getWriter();
                writer.println("<script language='javascript'>alert('파일을 찾을 수 없습니다');history.back();</script>");

            }
            in.close();
            os.close();
        } catch (Exception e) {
            System.out.println("ERROR : " + e.getMessage());
        }

    }

}