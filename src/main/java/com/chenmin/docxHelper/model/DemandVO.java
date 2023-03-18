package com.chenmin.docxHelper.model;

public class DemandVO {
    /**
     * 序号
     */
    private int order;

    /**
     * 需求编号
     */
    private String id;

    /**
     * 需求描述
     */
    private String description;

    public DemandVO() {
    }

    public DemandVO(int order, String id, String description) {
        this.order = order;
        this.id = id;
        this.description = description;
    }

    public int getOrder() {
        return order;
    }

    public void setOrder(int order) {
        this.order = order;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }
}
