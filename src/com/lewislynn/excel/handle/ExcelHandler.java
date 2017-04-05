/**
 *
 */
package com.lewislynn.excel.handle;

/**
 * <p>Title: com.lewislynn.excel.ExcelHandler.java</p>
 *
 * <p>Description: </p>
 *
 * <p>Copyright: Copyright (c) 2001-2013 Newland SoftWare Company</p>
 *
 * <p>Company: Newland SoftWare Company</p>
 *
 * @author Lewis.Lynn
 *
 * @version 1.0 CreateTime：2016年8月7日 下午10:03:37
 */

public interface ExcelHandler {

    public String readCell();

    public void writeCell(String value);
}
