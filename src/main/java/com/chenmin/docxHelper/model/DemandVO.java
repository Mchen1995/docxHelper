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

    /**
     * 案例个数
     */
    private int caseCount;

    /**
     * 姓名
     */
    private String name;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public DemandVO() {
    }

    public DemandVO(int order, String id, String description, int caseCount) {
        this.order = order;
        this.id = id;
        this.description = description;
        this.caseCount = caseCount;
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

    public int getCaseCount() {
        return caseCount;
    }

    public void setCaseCount(int caseCount) {
        this.caseCount = caseCount;
    }
}
