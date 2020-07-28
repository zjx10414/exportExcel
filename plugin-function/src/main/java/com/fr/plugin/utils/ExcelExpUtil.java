package com.fr.plugin.utils;

import org.apache.commons.codec.binary.Base64;
import org.apache.poi.ss.usermodel.Workbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;

public class ExcelExpUtil {

    public static void outputXls(Workbook workbook, String fileName, HttpServletResponse response, HttpServletRequest request) {
        ByteArrayOutputStream os = new ByteArrayOutputStream();

        try {
            workbook.write(os);
            byte[] content = os.toByteArray();
            InputStream is = new ByteArrayInputStream(content);
            // 设置response参数，可以打开下载页面
            response.reset();
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader("Content-Disposition", "attachment;filename="+encodeFileName(fileName,request));
            ServletOutputStream out = response.getOutputStream();
            BufferedInputStream bis = null;
            BufferedOutputStream bos = null;
            try {
                bis = new BufferedInputStream(is);
                bos = new BufferedOutputStream(out);
                byte[] buff = new byte[2048];
                int bytesRead;
                // Simple read/write loop.
                while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
                    bos.write(buff, 0, bytesRead);
                }
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                if (bis != null){
                    bis.close();
                }
                if (bos != null){
                    bos.close();
                }
            }
        } catch (Exception e1) {
            e1.printStackTrace();
        }
    }

    public static String encodeFileName(String fileName, HttpServletRequest request)
            throws UnsupportedEncodingException {
        String agent = request.getHeader("USER-AGENT");
        if (null != agent && -1 != agent.indexOf("MSIE")) {
            return URLEncoder.encode(fileName, "UTF-8");
        } else if (null != agent && -1 != agent.indexOf("Mozilla")) {
            return "=?UTF-8?B?"
                    + (new String(
                    Base64.encodeBase64(fileName.getBytes("UTF-8"))))
                    + "?=";
        } else {
            return fileName;
        }
    }

}
