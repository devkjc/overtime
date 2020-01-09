package com.fine.overtime.controller;

import lombok.extern.java.Log;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.*;

@Controller
@Log
@RequestMapping("common")
public class CommonDownLoadFile {
	
	@RequestMapping("/fileDown")
    private void fileDown(String fileUrl,
                          String oriFileName, HttpServletRequest request, HttpServletResponse response, HttpSession session) throws Exception{
        
        request.setCharacterEncoding("UTF-8");
        //System.out.println("파일이름: "+oriFileName);
        
        //파일 업로드된 경로 
        try{
        	
        	String defaultPth = session.getServletContext().getRealPath("/"); // 서버 기본 경로 (프로젝트 X)
            
        	//System.out.println("defaultPth:" + defaultPth);
        	fileUrl = defaultPth + File.separator + fileUrl;
        	
        	//System.out.println("fileUrl:" + fileUrl);
            		
            		
            //실제 내보낼 파일명 
            InputStream in = null;
            OutputStream os = null;
            File file = null;
            boolean skip = false;
            String client = "";
            
            //파일을 읽어 스트림에 담기  
            try{
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
            	String[] invalidName = {"\\\\","/",":","[*]","[?]","\"","<",">","[|]","&", " "}; // 윈도우 파일명으로 사용할수 없는 문자
                
                for(int i=0; i<invalidName.length; i++) {
                	oriFileName = oriFileName.replaceAll(invalidName[i], " ");
                }
                
                //System.out.println("변환된 파일이름: "+oriFileName);
                
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
            //System.out.println("ERROR : " + e.getMessage());
        }
        
    }

}
