package com.fine.overtime.util.excel.read;

import lombok.extern.java.Log;
import org.springframework.ui.Model;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpSession;
import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;
import java.util.UUID;

@Log
public class ExcelUploadUtil {
	
	public File execute(Model model) {
		
		Map<String, Object> map = model.asMap();
		
		MultipartFile mf = (MultipartFile) map.get("excelFile");
		HttpSession httpSession  = (HttpSession) map.get("httpSession");
		
		if(mf != null && !(mf.getOriginalFilename().equals(""))) {

			String originalName = mf.getOriginalFilename(); //파일이름
			String originalNameExtension = originalName.substring(originalName.lastIndexOf(".") + 1).toLowerCase();

			//확장자 제한
			if (((originalNameExtension.equals("html")) || (originalNameExtension.equals("php")) || (originalNameExtension.equals("exe")) || (originalNameExtension.equals("js")) || (originalNameExtension.equals("java")) || (originalNameExtension.equals("class")))) {
				return null;
			}

			//파일 크기 제한(100MB)
			long fileSize = mf.getSize();
			long limitFileSize = 100 * 1024 * 1024; //100MB
			if (limitFileSize < fileSize) {
				return null;
			}

			//파일 저장명 처리
			SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMddHHmmss");
			SimpleDateFormat dateFormatHH = new SimpleDateFormat("yyyyMM");
			String today = dateFormat.format(new Date());
			String todayHH = dateFormatHH.format(new Date());
			String modifyPath = todayHH;
			String modifyName = today + "_" + UUID.randomUUID().toString().substring(20) + "." + originalNameExtension;

			//저장경로
			String defaultPath = httpSession.getServletContext().getRealPath("/"); //서버 기본 경로(프로젝트X)
			String path = defaultPath + File.separator + "upload" + File.separator + "board" + File.separator + "file" + File.separator + modifyPath + File.separator;
			
			//저장경로 처리
			File file = new File(path);
			if(!file.exists()) { //디렉토리 존재하지 않을 경우 디렉토리 생성
				file.mkdirs();
			}
			
			//Multipart 처리
			try {
				File excel = new File(path + modifyName);
				mf.transferTo(excel);
				log.info("path + modifyName : " + path + modifyName);
				log.info("path : " + path);
				log.info("originalName : " + originalName);
				return excel;
			} catch(Exception e) {
				e.printStackTrace();
				return null;
			}
		}else {
			return null;
		}
	}

}
