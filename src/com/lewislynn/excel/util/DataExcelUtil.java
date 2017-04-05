/**
 *
 */
package com.lewislynn.excel.util;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.lewislynn.excel.model.DataModel;

import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.VerticalAlignment;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableImage;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * <p>Title: com.lewislynn.excel.util.ExcelUtil.java</p>
 *
 * <p>Description: </p>
 *
 * <p>Copyright: Copyright (c) 2001-2013 Newland SoftWare Company</p>
 *
 * <p>Company: Newland SoftWare Company</p>
 *
 * @author Lewis.Lynn
 *
 * @version 1.0 CreateTime��2016��8��7�� ����10:01:54
 */

public class DataExcelUtil {

    private Demo1 demo1;

    public static List<DataModel> getAllData(){

        List<DataModel> models = null;
        try {
            models = new ArrayList<DataModel>();

            String targetExcel = "C:\\Users\\Lewis\\Desktop\\做标签\\new20160818.xls";

            Workbook wb = Workbook.getWorkbook(new File(targetExcel));
            Sheet sheet = wb.getSheet("詹斌斌");

            for (int row = 1; row < 201; row++) {
                DataModel model = new DataModel();
                model.setName(sheet.getCell(0, row).getContents());
                model.setId(sheet.getCell(1, row).getContents());
                model.setModel(sheet.getCell(2, row).getContents());
                model.setPrice(Double.parseDouble(sheet.getCell(3, row).getContents()));
                model.setDatetime(sheet.getCell(4, row).getContents());
                model.setPlace(sheet.getCell(5, row).getContents());
                model.setUsedDepartment(sheet.getCell(6, row).getContents());
                model.setUsername(sheet.getCell(7, row).getContents());
                models.add(model);
            }

            wb.close();
        } catch (NumberFormatException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return models;
    }

    public void setDetail(Demo2 d2){

    }

    public static void main(String[] args) {
        List<DataModel> models = DataExcelUtil.getAllData();
        System.out.println(" models size is " + models.size());
    }
}
