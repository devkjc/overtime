package com.fine.overtime.controller;

import com.fine.overtime.util.FileUtils;
import lombok.RequiredArgsConstructor;
import lombok.extern.java.Log;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.File;

@Controller
@Log
@RequestMapping("common")
@RequiredArgsConstructor
public class CommonDownLoadFile {

    private final FileUtils fileUtils;

    @RequestMapping("/fileDown")
    private void fileDown(String fileUrl, String fileName,
                          String oriFileName, HttpServletRequest request, HttpServletResponse response, HttpSession session) throws Exception {

        String defaultPth = session.getServletContext().getRealPath("/");
        fileUrl = defaultPth + File.separator + fileUrl;

        fileUtils.fileDown(fileUrl, fileName, oriFileName, request, response, session);
    }
}
