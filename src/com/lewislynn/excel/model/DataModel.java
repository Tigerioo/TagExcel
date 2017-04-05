/**
 *
 */
package com.lewislynn.excel.model;

/**
 * <p>Title: com.lewislynn.excel.model.DataModel.java</p>
 *
 * <p>Description: </p>
 *
 * <p>Copyright: Copyright (c) 2001-2013 Newland SoftWare Company</p>
 *
 * <p>Company: Newland SoftWare Company</p>
 *
 * @author Lewis.Lynn
 *
 * @version 1.0 CreateTime：2016年8月7日 下午11:54:11
 */

public class DataModel {

    private String name;//资产名称
    private String id;//资产编号
    private String model;//资产型号
    private double price;//资产原值
    private String datetime;//购买日期
    private String place;//存放日期
    private String usedDepartment;//使用部门
    private String username;//责任人
    public String getName() {
        return name;
    }
    public void setName(String name) {
        this.name = name;
    }
    public String getId() {
        return id;
    }
    public void setId(String id) {
        this.id = id;
    }
    public String getModel() {
        return model;
    }
    public void setModel(String model) {
        this.model = model;
    }
    public double getPrice() {
        return price;
    }
    public void setPrice(double price) {
        this.price = price;
    }
    public String getDatetime() {
        return datetime;
    }
    public void setDatetime(String datetime) {
        this.datetime = datetime;
    }
    public String getPlace() {
        return place;
    }
    public void setPlace(String place) {
        this.place = place;
    }
    public String getUsedDepartment() {
        return usedDepartment;
    }
    public void setUsedDepartment(String usedDepartment) {
        this.usedDepartment = usedDepartment;
    }
    public String getUsername() {
        return username;
    }
    public void setUsername(String username) {
        this.username = username;
    }

}
