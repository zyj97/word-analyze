package org.analyze.pojo;

import org.springframework.stereotype.Component;

/**
 * @author zhanlin.mao
 * @date 5/7/21 5:33 PM
 */

@Component
public class TableContent {
    private String id;
    private String content;
    private Integer order;
    private Integer rowId;
    private Integer columnId;
    private String tabNo;
    private String title;
    //表格宽度 用于
    private int tableWidth;

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getContent() {
        return content;
    }

    public void setContent(String content) {
        this.content = content;
    }

    public Integer getOrder() {
        return order;
    }

    public void setOrder(Integer order) {
        this.order = order;
    }

    public Integer getRowId() {
        return rowId;
    }

    public void setRowId(Integer rowId) {
        this.rowId = rowId;
    }

    public Integer getColumnId() {
        return columnId;
    }

    public void setColumnId(Integer columnId) {
        this.columnId = columnId;
    }

    public String getTabNo() {
        return tabNo;
    }

    public void setTabNo(String tabNo) {
        this.tabNo = tabNo;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public int getTableWidth() {
        return tableWidth;
    }

    public void setTableWidth(int tableWidth) {
        this.tableWidth = tableWidth;
    }

    @Override
    public String toString() {
        return "TableContent{" +
                "id='" + id + '\'' +
                ", content='" + content + '\'' +
                ", order=" + order +
                ", rowId=" + rowId +
                ", columnId=" + columnId +
                ", tabNo='" + tabNo + '\'' +
                ", title='" + title + '\'' +
                ", tableWidth=" + tableWidth +
                '}';
    }
}
