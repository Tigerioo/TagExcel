/**
 *
 */
package com.lewislynn.excel.util;

import java.io.File;
import java.io.IOException;
import java.util.List;

import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.NumberFormat;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableImage;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import com.lewislynn.excel.model.DataModel;

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
 * @version 1.0 CreateTime 2016.8.7 10:01:54
 */

public class ExcelUtil {

    private WritableWorkbook targetWorkbook;
    private WritableSheet mySheet;

    public ExcelUtil(){
        String targetExcel = "C:\\Users\\Lewis\\Desktop\\做标签\\20160818\\詹斌斌.xls";

        try {
            targetWorkbook = Workbook.createWorkbook(new File(targetExcel));
            mySheet = targetWorkbook.createSheet("Sheet1", 0);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * set title
     * @param r	row
     * @param c	column
     * @throws WriteException
     */
    public void createTitle(int r, int c) throws WriteException {

        WritableFont font1 = new WritableFont(WritableFont.createFont("宋体"), 10);
        WritableCellFormat format = new WritableCellFormat(font1);
        //设置水平居中
        format.setAlignment(Alignment.CENTRE);
        //设置垂直居中
        format.setVerticalAlignment(VerticalAlignment.CENTRE);
        Label label = new Label(r, c, "连江恒欣村镇银行固定资产", format);
        mySheet.addCell(label);

        //set row height
        mySheet.setRowView(c, 435);

        //merge cell
        mySheet.mergeCells(r, c, r+3, c);

        //image file
        File imageFile = new File("C:\\Users\\Lewis\\Desktop\\做标签\\hengxin.png");
        WritableImage image = new WritableImage(r + 3.3, c + 0.2, 0.34, 0.60, imageFile);
        mySheet.addImage(image);
    }

    /**
     *
     * @param c cloumn
     * @param r row
     * @throws WriteException
     */
    public void createContent(int c, int r, WritableCellFormat format) throws WriteException{
        Label name = new Label(c, r, "资产名称", format);
        mySheet.addCell(name);

        Label id = new Label(c+2, r, "资产编号", format);
        mySheet.addCell(id);

        Label model = new Label(c, r+1, "规格型号", format);
        mySheet.addCell(model);

        Label price = new Label(c+2, r+1, "资产原值", format);
        mySheet.addCell(price);

        Label dateTime = new Label(c, r+2, "购置日期", format);
        mySheet.addCell(dateTime);

        Label place = new Label(c+2, r+2, "存放地点", format);
        mySheet.addCell(place);

        Label usedDepartment = new Label(c, r+3, "使用部门", format);
        mySheet.addCell(usedDepartment);

        Label user = new Label(c+2, r+3, "负责人", format);
        mySheet.addCell(user);
    }

    /**
     *
     * @param c cloumn
     * @param r row
     * @throws WriteException
     */
    public void setValue(int c, int r, DataModel model, WritableCellFormat format, WritableCellFormat numberCellFormat) throws WriteException{

        Label nameValue = new Label(c, r, model.getName(), format);
        mySheet.addCell(nameValue);

        Label idValue = new Label(c+2, r, model.getId(), format);
        mySheet.addCell(idValue);

        Label modelValue = new Label(c, r+1, model.getModel(), format);
        mySheet.addCell(modelValue);

        Number priceValue = new Number(c+2, r+1, model.getPrice(), numberCellFormat);
        mySheet.addCell(priceValue);

        Label dateTimeValue = new Label(c, r+2, model.getDatetime(), format);
        mySheet.addCell(dateTimeValue);

        Label placeValue = new Label(c+2, r+2, model.getPlace(), format);
        mySheet.addCell(placeValue);

        Label usedDepartmentValue = new Label(c, r+3, model.getUsedDepartment(), format);
        mySheet.addCell(usedDepartmentValue);

        Label userValue = new Label(c+2, r+3, model.getUsername(), format);
        mySheet.addCell(userValue);
    }

    public void close(){
        try {
            targetWorkbook.write();
            targetWorkbook.close();
        } catch (WriteException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public WritableCellFormat getCommonFormat(){
        WritableCellFormat format = null;
        try {
            WritableFont font1 = new WritableFont(WritableFont.createFont("宋体"), 9);
            format = new WritableCellFormat(font1);
            format.setAlignment(Alignment.CENTRE);
            format.setVerticalAlignment(VerticalAlignment.CENTRE);
            format.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
        } catch (WriteException e1) {
            e1.printStackTrace();
        }
        return format;
    }

    public WritableCellFormat getNumberFormat(){
        WritableCellFormat numberCellFormat = null;
        try {
            NumberFormat fivedps = new NumberFormat("#.00");
            numberCellFormat = new WritableCellFormat(fivedps);
            numberCellFormat.setAlignment(Alignment.CENTRE);
            numberCellFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
            numberCellFormat.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
        } catch (WriteException e) {
            e.printStackTrace();
        }
        return numberCellFormat;
    }

    public static void main(String[] args) {
        ExcelUtil util = new ExcelUtil();
        WritableCellFormat format = util.getCommonFormat();
        WritableCellFormat numberCellFormat = util.getNumberFormat();
        try {
            List<DataModel> models = DataExcelUtil.getAllData();

            for (int i = 0; i < models.size()*2; i+=6) {

                util.createTitle(0, i);
                util.createTitle(5, i);
                util.createTitle(10, i);

                util.createContent(0, i+1, format);
                util.createContent(5, i+1, format);
                util.createContent(10, i+1, format);
            }
            DataModel maxPriceModel = null;
            for (int i = 0; i < models.size(); i++) {
                DataModel model = models.get(i);
                if(maxPriceModel == null){
                    maxPriceModel = model;
                }else {
                    if(model.getPrice() > maxPriceModel.getPrice()){
                        maxPriceModel = model;
                    }
                }
                if(i % 3 == 0) {
                    System.out.println("i = " + i);
                    util.setValue(1, 2*i+1, model, format, numberCellFormat);
                }else if(i % 3 == 1){
                    util.setValue(6, 2*i-1, model, format, numberCellFormat);
                }else if(i % 3 == 2){
                    util.setValue(11, 2*i-3, model, format, numberCellFormat);
                }
            }
            if(maxPriceModel != null){
                System.out.println(maxPriceModel.getName() + ":" + maxPriceModel.getPrice());
            }

        } catch (WriteException e) {
            e.printStackTrace();
        } finally {
            util.close();
        }

    }
}
