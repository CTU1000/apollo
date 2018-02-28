/**
 * Alipay.com Inc.
 * Copyright (c) 2004-2018 All Rights Reserved.
 */
package com.ctrip.framework.apollo.portal.sheet;

import com.google.common.collect.Lists;

import java.util.List;

/**
 *
 * @author lepdou
 * @version $Id: Sheet.java, v 0.1 2018年02月28日 下午7:57 lepdou Exp $
 */
public class Sheet {

    private String       unit;
    private String       sheetName;
    private List<Double> data;

    public Sheet(String unit, String sheetName) {
        this.unit = unit;
        this.sheetName = sheetName;
        data = Lists.newArrayList();
    }

    public void addData(Double value) {
        data.add(value);
    }

    public String getUnit() {
        return unit;
    }

    public void setUnit(String unit) {
        this.unit = unit;
    }

    public Double getData(int index) {
        return data.get(index);
    }

    public List<Double> getData() {
        return data;
    }

    public String getSheetName() {
        return sheetName;
    }
}