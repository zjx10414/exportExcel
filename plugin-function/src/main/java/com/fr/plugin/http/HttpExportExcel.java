package com.fr.plugin.http;

import com.fanruan.api.data.ConnectionKit;
import com.fr.data.impl.AbstractDBDataModel;
import com.fr.data.impl.Connection;
import com.fr.data.impl.DBTableData;
import com.fr.decision.fun.impl.BaseHttpHandler;
import com.fr.plugin.entity.ExcelEntity;
import com.fr.plugin.utils.ExcelExpUtil;
import com.fr.third.springframework.web.bind.annotation.RequestMethod;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;
/**
 * 计算部分
 */
public class HttpExportExcel extends BaseHttpHandler {

    /**
     * GET,POST,DELETE
     * @return
     */
    @Override
    public RequestMethod getMethod() {
        return RequestMethod.GET;
    }

    /**
     *
     * @return
     */
    @Override
    public String getPath() {
        return "/export";
    }

    @Override
    public boolean isPublic() {
        return true;
    }

    @Override
    public void handle(HttpServletRequest request, HttpServletResponse response) throws Exception {

        XSSFWorkbook wb = new XSSFWorkbook();//创建工作薄
        XSSFCellStyle titleStyle = getTitleCellStyle(wb);
        XSSFCellStyle textStyle = getTextCellStyle(wb, HorizontalAlignment.CENTER);

        //数据库连接名
        String connectName = request.getParameter("connectName");
        //字典表表名
        String dictTable = request.getParameter("dictTable");
        //导出表表名集合
        String exportTables = request.getParameter("exportTables");
        //受试者id
        String subjectids = request.getParameter("subjectids");

        //获取数据库连接
        Connection connection = ConnectionKit.getConnection(connectName);
        //查询字典表
        String dictSql = "select * from " + dictTable;
        AbstractDBDataModel dictData = DBTableData.createCacheableDBResultSet(connection, dictSql, -1);
        Map<String,Object> dictMap = new HashMap<>();
        for (int row = 0; row < dictData.getRowCount(); row ++){
            String dreport = dictData.getValueAt(row, 1).toString();
            String edreport = dictData.getValueAt(row, 2).toString();
            String fields = dictData.getValueAt(row, 3).toString();
            ExcelEntity entity = new ExcelEntity();
            entity.setDreport(dreport);
            entity.setEdreport(edreport);
            entity.setFields(fields);
            dictMap.put(edreport,entity);
        }

        //查询导出数据
        List<String> exportTableList =  Arrays.asList(exportTables.split(";"));
        exportTableList = exportTableList.stream().filter(t-> !t.isEmpty()).collect(Collectors.toList());
        List<String> subjectidList = Arrays.asList(subjectids.split(";"));
        subjectidList = subjectidList.stream().filter(t-> !t.isEmpty()).collect(Collectors.toList());
        StringBuffer buf = new StringBuffer();
        for (int i = 0; i < subjectidList.size(); i++) {
            buf.append(subjectidList.get(i));
            if(i < subjectidList.size()-1){
                buf.append("','");
            }
        }
        String subjectidStr = buf.toString();
        for (String exportTable : exportTableList) {
            String sqlStr = "select * from " + exportTable + " where subjectid in ( '" + subjectidStr +" ')";
            AbstractDBDataModel dataModel = DBTableData.createCacheableDBResultSet(connection, sqlStr, -1);
            Integer rowCount = dataModel.getRowCount();
            Integer columnCount = dataModel.getColumnCount();

            ExcelEntity outExcel =(ExcelEntity)dictMap.get(exportTable);
            List<String> titleList = Arrays.asList(outExcel.getFields().split(";"));
            titleList = titleList.stream().filter(t-> !t.isEmpty()).collect(Collectors.toList());

            //新建一个sheet,设置sheet名,设置列宽
            XSSFSheet sheet = wb.createSheet(outExcel.getDreport());
            for (int i = 0; i < columnCount; i++) {
                sheet.setColumnWidth(i,256 * 20 );
            }

            //标题行
            XSSFRow titleRow = sheet.createRow(0);
            titleRow.setHeight((short)600);
            for(int i=0 ;i < titleList.size() ; i++){
                XSSFCell cell = titleRow.createCell(i);
                //往单元格里写数据
                cell.setCellValue(titleList.get(i));
                cell.setCellStyle(titleStyle); //加样式
            }

            //写数据集
            for (int row = 0; row < rowCount; row ++){
                XSSFRow textRow = sheet.createRow(row + 1);
                textRow.setHeight((short)600);
                for (int colum = 0 ; colum < columnCount; colum ++){
                    XSSFCell cell = textRow.createCell(colum);
                    // 数据值
                    String value = dataModel.getValueAt(row, colum).toString();
                    cell.setCellValue(value);
                    cell.setCellStyle(textStyle); //加样式
                }
            }
        }
        System.out.println("==========================编译成功");
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");
        String dateNowStr = sdf.format(new Date());
        String fileName ="受试者信息_"+dateNowStr+".xlsx";
        ExcelExpUtil.outputXls(wb, fileName, response, request);
    }

    private static XSSFCellStyle getTitleCellStyle(XSSFWorkbook wb) {
        XSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);//水平居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
        XSSFFont fontStyle = wb.createFont();
        fontStyle.setFontName("宋体"); //字体样式
        fontStyle.setFontHeightInPoints((short)10);
        fontStyle.setBold(true);//加粗
        style.setFont(fontStyle);
        style.setBorderBottom(BorderStyle.THIN); //下边框
        style.setBorderLeft(BorderStyle.THIN);//左边框
        style.setBorderTop(BorderStyle.THIN);//上边框
        style.setBorderRight(BorderStyle.THIN);//右边框
        return style;
    }

    private static XSSFCellStyle getTextCellStyle(XSSFWorkbook wb, HorizontalAlignment setAlignment) {
        XSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(setAlignment);//水平居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
        style.setWrapText(true);
        XSSFFont fontStyle = wb.createFont();
        fontStyle.setFontName("宋体"); //字体样式
        fontStyle.setFontHeightInPoints((short)8);
        style.setFont(fontStyle);
        style.setBorderBottom(BorderStyle.THIN); //下边框
        style.setBorderLeft(BorderStyle.THIN);//左边框
        style.setBorderTop(BorderStyle.THIN);//上边框
        style.setBorderRight(BorderStyle.THIN);//右边框
        return style;
    }
}
