package com.ewancle.service;

import com.ewancle.model.Employee;
import io.smallrye.mutiny.Multi;
import io.smallrye.mutiny.Uni;
import io.vertx.core.buffer.Buffer;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import jakarta.enterprise.context.ApplicationScoped;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

@ApplicationScoped
public class ExcelService {

    // 模拟数据源 - 实际项目中可能来自数据库
    public Multi<Employee> getEmployeeStream() {
        List<Employee> employees = Arrays.asList(
                new Employee(1L, "张三", "技术部", 15000.0, "zhangsan@company.com"),
                new Employee(2L, "李四", "产品部", 12000.0, "lisi@company.com"),
                new Employee(3L, "王五", "市场部", 11000.0, "wangwu@company.com"),
                new Employee(4L, "赵六", "人事部", 10000.0, "zhaoliu@company.com"),
                new Employee(5L, "钱七", "财务部", 13000.0, "qianqi@company.com"),
                new Employee(6L, "孙八", "技术部", 16000.0, "sunba@company.com"),
                new Employee(7L, "周九", "产品部", 14000.0, "zhoujiu@company.com"),
                new Employee(8L, "吴十", "市场部", 9000.0, "wushi@company.com")
        );

        // 模拟流式数据，每500ms发送一条记录
        return Multi.createFrom().iterable(employees)
                .onItem().call(item -> Uni.createFrom().nullItem()
                        .onItem().delayIt().by(Duration.ofMillis(500)));
    }

    // 生成Excel的流式Buffer数据 - 修复版本
    public Multi<Buffer> generateExcelStreamAsBuffer() {
        return getEmployeeStream()
                .collect().asList()
                .toMulti()
                .flatMap(employees -> {
                    return Multi.createFrom().emitter(emitter -> {
                        try (SXSSFWorkbook workbook = new SXSSFWorkbook(1000)){
                            // 使用SXSSFWorkbook进行流式处理，减少内存占用
                            Sheet sheet = workbook.createSheet("员工信息");

                            // 创建样式
                            CellStyle headerStyle = createHeaderStyle(workbook);
                            CellStyle dataStyle = createDataStyle(workbook);

                            // 创建标题行
                            createHeaderRow(sheet, headerStyle);


                            // 分批处理数据，每批发送一个Buffer
                            int batchSize = 100; // 每100行发送一次
                            for (int i = 0; i < employees.size(); i += batchSize) {
                                int endIndex = Math.min(i + batchSize, employees.size());
                                List<Employee> batch = employees.subList(i, endIndex);

                                // 添加当前批次的数据
                                fillDataRowsBatch(sheet, batch, dataStyle, i + 1);

                                // 创建新的输出流并写入当前状态
                                try (ByteArrayOutputStream batchOut = new ByteArrayOutputStream();
                                     SXSSFWorkbook tempWorkbook = new SXSSFWorkbook(1000);// 创建临时工作簿副本用于输出
                                ) {
                                    Sheet tempSheet = tempWorkbook.createSheet("员工信息");

                                    // 复制样式
                                    CellStyle tempHeaderStyle = createHeaderStyle(tempWorkbook);
                                    CellStyle tempDataStyle = createDataStyle(tempWorkbook);

                                    // 复制表头
                                    createHeaderRow(tempSheet, tempHeaderStyle);

                                    // 复制到当前为止的所有数据
                                    fillDataRows(tempSheet, employees.subList(0, endIndex), tempDataStyle);

                                    /*if( i == 0){
                                        // 自动调整列宽
                                        for (int col = 0; col < 5; col++) {
                                            tempSheet.autoSizeColumn(col+1);
                                        }
                                    }*/

                                    tempWorkbook.write(batchOut);

                                    Buffer buffer = Buffer.buffer(batchOut.toByteArray());
                                    emitter.emit(buffer);

                                } catch (IOException e) {
                                    emitter.fail(e);
                                    return;
                                }
                            }

                            emitter.complete();

                        } catch (Exception e) {
                            emitter.fail(e);
                        }
                    });
                });
    }

    // 简化版本：直接生成完整Excel作为流
    public Multi<Buffer> generateExcelStreamAsBufferSimple() {
        return generateCompleteExcelAsBuffer().toMulti();
    }

    // 分批填充数据行
    private void fillDataRowsBatch(Sheet sheet, List<Employee> employees, CellStyle dataStyle, int startRowNum) {
        int currentRowNum = startRowNum;
        for (Employee employee : employees) {
            Row row = sheet.createRow(currentRowNum++);

            Cell cell0 = row.createCell(0);
            cell0.setCellValue(employee.getId());
            cell0.setCellStyle(dataStyle);

            Cell cell1 = row.createCell(1);
            cell1.setCellValue(employee.getName());
            cell1.setCellStyle(dataStyle);

            Cell cell2 = row.createCell(2);
            cell2.setCellValue(employee.getDepartment());
            cell2.setCellStyle(dataStyle);

            Cell cell3 = row.createCell(3);
            cell3.setCellValue(employee.getSalary());
            cell3.setCellStyle(dataStyle);

            Cell cell4 = row.createCell(4);
            cell4.setCellValue(employee.getEmail());
            cell4.setCellStyle(dataStyle);
        }
    }

    // 生成Excel的流式字节数据
    public Multi<byte[]> generateExcelStream() {
        return generateExcelStreamAsBuffer()
                .map(Buffer::getBytes);
    }

    // 优化版本：直接生成最终Excel文件为Buffer
    public Uni<Buffer> generateCompleteExcelAsBuffer() {
        return getEmployeeStream()
                .collect().asList()
                .map(employees -> {
                    try {
                        SXSSFWorkbook workbook = new SXSSFWorkbook(1000);
                        Sheet sheet = workbook.createSheet("员工信息");

                        // 创建样式
                        CellStyle headerStyle = createHeaderStyle(workbook);
                        CellStyle dataStyle = createDataStyle(workbook);

                        // 创建标题行
                        createHeaderRow(sheet, headerStyle);

                        // 填充数据
                        fillDataRows(sheet, employees, dataStyle);

                        // 自动调整列宽
                        for (int i = 0; i < 5; i++) {
                            sheet.autoSizeColumn(i);
                        }

                        ByteArrayOutputStream out = new ByteArrayOutputStream();
                        workbook.write(out);
                        workbook.close();

                        return Buffer.buffer(out.toByteArray());
                    } catch (IOException e) {
                        throw new RuntimeException("生成Excel文件失败", e);
                    }
                });
    }

    // 优化版本：直接生成最终Excel文件
    public Uni<byte[]> generateCompleteExcel() {
        return generateCompleteExcelAsBuffer()
                .map(buffer -> buffer.getBytes());
    }

    private CellStyle createHeaderStyle(SXSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        setBorders(style);
        return style;
    }

    private CellStyle createDataStyle(SXSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        setBorders(style);
        return style;
    }

    private void setBorders(CellStyle style) {
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
    }

    private void createHeaderRow(Sheet sheet, CellStyle headerStyle) {
        Row headerRow = sheet.createRow(0);
        String[] headers = {"ID", "姓名", "部门", "薪资", "邮箱"};
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }
    }

    private void fillDataRows(Sheet sheet, List<Employee> employees, CellStyle dataStyle) {
        int rowNum = 1;
        for (Employee employee : employees) {
            Row row = sheet.createRow(rowNum++);

            Cell cell0 = row.createCell(0);
            cell0.setCellValue(employee.getId());
            cell0.setCellStyle(dataStyle);

            Cell cell1 = row.createCell(1);
            cell1.setCellValue(employee.getName());
            cell1.setCellStyle(dataStyle);

            Cell cell2 = row.createCell(2);
            cell2.setCellValue(employee.getDepartment());
            cell2.setCellStyle(dataStyle);

            Cell cell3 = row.createCell(3);
            cell3.setCellValue(employee.getSalary());
            cell3.setCellStyle(dataStyle);

            Cell cell4 = row.createCell(4);
            cell4.setCellValue(employee.getEmail());
            cell4.setCellStyle(dataStyle);
        }
    }
}