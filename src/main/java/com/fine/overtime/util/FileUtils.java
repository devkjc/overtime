package com.fine.overtime.util;

import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Service
public class FileUtils {

    /**
     * 디렉토리 및 파일을 압축한다.
     *
     * @param path     압축할 디렉토리 및 파일
     * @param fileName 압축파일의 이름
     */

    public void createZipFile(String path, String fileName) {

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

            for (String s : list) {
                zip_folder("", new File(_path + s), zip_out, path);
            }

            zip_out.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
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

    public void zip_folder(String parent, File file, ZipOutputStream zout, String toPath) throws IOException {
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
                for (String s : list) {
                    zip_folder(entry.getName(), new File(file.getPath() + "/" + s), zout, toPath);
                }
            }
        }
    }

    public void directoryDelete(String path) {
        File folder = new File(path);
        try {
            while (folder.exists()) {
                File[] folder_list = folder.listFiles(); //파일리스트 얻어오기

                for (File file : folder_list) {
                    file.delete(); //파일 삭제
                }
                if (folder_list.length == 0 && folder.isDirectory()) {
                    folder.delete(); //대상폴더 삭제
                }
            }
        } catch (Exception e) {
            e.getStackTrace();
        }
    }

    public void fileDown(String fileUrl, String fileName, String oriFileName, HttpServletRequest request, HttpServletResponse response, HttpSession session) throws Exception {

        request.setCharacterEncoding("UTF-8");

        //파일 업로드된 경 로
        try {

            fileUrl = fileUrl + fileName;

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

                for (String s : invalidName) {
                    oriFileName = oriFileName.replaceAll(s, "_");
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
