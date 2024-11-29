package com.heartsuit.db2word.postgresql.service.impl;

import com.heartsuit.db2word.postgresql.dao.DataSourceMapper;
import com.heartsuit.db2word.postgresql.service.DataSourceDetailService;
import com.lowagie.text.*;
import com.lowagie.text.Font;
import com.lowagie.text.rtf.RtfWriter2;
import org.springframework.stereotype.Service;

import java.awt.*;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;

/**
 * Author:  Heartsuit
 * Date:  2021/6/8 8:33
 */
@Service
public class DataSourceDetailServiceImpl implements DataSourceDetailService {

    private final DataSourceMapper dataSourceMapper;

    public DataSourceDetailServiceImpl(DataSourceMapper dataSourceMapper) {
        this.dataSourceMapper = dataSourceMapper;
    }

    @Override
    public List<Map<String, Object>> getAllTableNames() {
        return dataSourceMapper.getAllTableNames();
    }

    @Override
    public List<Map<String, Object>> getTableColumnDetail(String tableName) {
        return dataSourceMapper.getTableColumnDetail(tableName);
    }

    @Override
    public void toWord(List<Map<String, Object>> tables) throws FileNotFoundException, DocumentException {
        // 创建word文档,并设置纸张的大小
        Document document = new Document(PageSize.A4);
        // 创建word文档
        RtfWriter2.getInstance(document, new FileOutputStream("C:/Users/wendy/Desktop/word/zq_onemap.doc"));
        document.open();// 设置文档标题
        Paragraph p = new Paragraph("数据库表设计文档", new Font(Font.NORMAL, 24, Font.BOLD, new Color(0, 0, 0)));
        p.setAlignment(1);
        document.add(p);

        /* * 创建表格 通过查询出来的表遍历 */
        for (int i = 0; i < tables.size(); i++) {
            // 表名
            String table_name = (String) tables.get(i).get("table_name");
            // 表说明
            String table_comment = tables.get(i).get("table_comment") == null ? "" : (String) tables.get(i).get("table_comment");

            //获取某张表的所有字段说明
            List<Map<String, Object>> columns = this.getTableColumnDetail(table_name);
            //构建表说明
            String all = table_comment + "表(" + table_name + ")";
            //创建有6列的表格
            Table table = new Table(6);
            document.add(new Paragraph(""));
            table.setBorderWidth(1);
            table.setPadding(0);
            table.setSpacing(0);

            /*
             * 添加表头的元素
             */
            // 创建表头字体（加粗）
            Font boldFont = new Font(Font.HELVETICA, 12, Font.BOLD);
            Cell cell = new Cell(new Chunk("序号", boldFont));// 单元格
            cell.setHeader(true);
            cell.setBackgroundColor(new Color(224, 224, 224));
            cell.setHorizontalAlignment(Cell.ALIGN_CENTER);
            table.addCell(cell);

            cell = new Cell(new Chunk("名称", boldFont));// 单元格
            cell.setHorizontalAlignment(Cell.ALIGN_CENTER);
            cell.setBackgroundColor(new Color(224, 224, 224));
            table.addCell(cell);

            cell = new Cell(new Chunk("类型", boldFont));// 单元格
            cell.setHorizontalAlignment(Cell.ALIGN_CENTER);
            cell.setBackgroundColor(new Color(224, 224, 224));
            table.addCell(cell);

            cell = new Cell(new Chunk("允许为空", boldFont));// 单元格
            cell.setHorizontalAlignment(Cell.ALIGN_CENTER);
            cell.setBackgroundColor(new Color(224, 224, 224));
            table.addCell(cell);

            cell = new Cell(new Chunk("索引", boldFont));// 单元格
            cell.setHorizontalAlignment(Cell.ALIGN_CENTER);
            cell.setBackgroundColor(new Color(224, 224, 224));
            table.addCell(cell);

            cell = new Cell(new Chunk("说明", boldFont));// 单元格
            cell.setHorizontalAlignment(Cell.ALIGN_CENTER);
            cell.setBackgroundColor(new Color(224, 224, 224));
            table.addCell(cell);

            table.endHeaders();// 表头结束

            // 表格的主体
            for (int k = 0; k < columns.size(); k++) {
                //获取某表每个字段的详细说明
                String Field = (String) columns.get(k).get("field");
                String Type = (String) columns.get(k).get("type");
                String Null = (String) columns.get(k).get("null");
                String Key = (String) columns.get(k).get("key");
                String Comment = (String) columns.get(k).get("comment");
                // 创建并设置单元格
                Cell indexCell = new Cell((k + 1) + ""); // 索引列
                indexCell.setHorizontalAlignment(Cell.ALIGN_CENTER);
                indexCell.setVerticalAlignment(Cell.ALIGN_MIDDLE);
                table.addCell(indexCell);

                Cell fieldCell = new Cell(Field); // 字段列
                fieldCell.setHorizontalAlignment(Cell.ALIGN_CENTER);
                fieldCell.setVerticalAlignment(Cell.ALIGN_MIDDLE);
                table.addCell(fieldCell);

                Cell typeCell = new Cell(Type); // 类型列
                typeCell.setHorizontalAlignment(Cell.ALIGN_CENTER);
                typeCell.setVerticalAlignment(Cell.ALIGN_MIDDLE);
                table.addCell(typeCell);

                Cell nullCell = new Cell(Null); // 是否为空列
                nullCell.setHorizontalAlignment(Cell.ALIGN_CENTER);
                nullCell.setVerticalAlignment(Cell.ALIGN_MIDDLE);
                table.addCell(nullCell);

                Cell keyCell = new Cell(Key); // 主键列
                keyCell.setHorizontalAlignment(Cell.ALIGN_CENTER);
                keyCell.setVerticalAlignment(Cell.ALIGN_MIDDLE);
                table.addCell(keyCell);

                Cell commentCell = new Cell(Comment); // 注释列
                commentCell.setHorizontalAlignment(Cell.ALIGN_CENTER);
                commentCell.setVerticalAlignment(Cell.ALIGN_MIDDLE);
                table.addCell(commentCell);
            }
            Paragraph paragraph = new Paragraph(all, new Font(Font.NORMAL, 14, Font.BOLD, new Color(0, 0, 0)));
            //写入表说明
            document.add(paragraph);
            //生成表格
            document.add(table);
        }
        document.close();
    }

    @Override
    public void toWordAllTable(List<Map<String, Object>> tables) throws FileNotFoundException, DocumentException {
        // 创建word文档,并设置纸张的大小
        Document document = new Document(PageSize.A4);
        // 创建word文档
        RtfWriter2.getInstance(document, new FileOutputStream("C:/Users/wendy/Desktop/word/zq_bpm-allTable.doc"));
        document.open();// 设置文档标题
        Paragraph p = new Paragraph("数据库表设计文档", new Font(Font.NORMAL, 24, Font.BOLD, new Color(0, 0, 0)));
        p.setAlignment(1);
        document.add(p);

        //创建有6列的表格
        Table table = new Table(3);
        document.add(new Paragraph(""));
        table.setBorderWidth(1);
        table.setPadding(0);
        table.setSpacing(0);

        /*
         * 添加表头的元素
         */
        // 创建表头字体（加粗）
        Font boldFont = new Font(Font.HELVETICA, 12, Font.BOLD);
        Cell cell = new Cell(new Chunk("序号", boldFont));// 单元格
        cell.setHeader(true);
        cell.setBackgroundColor(new Color(224, 224, 224));
        cell.setHorizontalAlignment(Cell.ALIGN_CENTER);
        table.addCell(cell);

        cell = new Cell(new Chunk("表名", boldFont));// 单元格
        cell.setHorizontalAlignment(Cell.ALIGN_CENTER);
        cell.setBackgroundColor(new Color(224, 224, 224));
        table.addCell(cell);

        cell = new Cell(new Chunk("注释", boldFont));// 单元格
        cell.setHorizontalAlignment(Cell.ALIGN_CENTER);
        cell.setBackgroundColor(new Color(224, 224, 224));
        table.addCell(cell);

        table.endHeaders();// 表头结束

        /* * 创建表格 通过查询出来的表遍历 */
        for (int i = 0; i < tables.size(); i++) {
            // 表名
            String table_name = (String) tables.get(i).get("table_name");
            // 表说明
            String table_comment = tables.get(i).get("table_comment") == null ? "" : (String) tables.get(i).get("table_comment");

            // 创建并设置单元格
            Cell indexCell = new Cell((i + 1) + ""); // 索引列
            indexCell.setHorizontalAlignment(Cell.ALIGN_CENTER);
            indexCell.setVerticalAlignment(Cell.ALIGN_MIDDLE);
            table.addCell(indexCell);

            Cell fieldCell = new Cell(table_name); // 字段列
            fieldCell.setHorizontalAlignment(Cell.ALIGN_CENTER);
            fieldCell.setVerticalAlignment(Cell.ALIGN_MIDDLE);
            table.addCell(fieldCell);

            Cell commentCell = new Cell(table_comment); // 注释列
            commentCell.setHorizontalAlignment(Cell.ALIGN_CENTER);
            commentCell.setVerticalAlignment(Cell.ALIGN_MIDDLE);
            table.addCell(commentCell);
        }
//            Paragraph paragraph = new Paragraph("", new Font(Font.NORMAL, 14, Font.BOLD, new Color(0, 0, 0)));
//            //写入表说明
//            document.add(paragraph);
        //生成表格
        document.add(table);
        document.close();
    }
}
