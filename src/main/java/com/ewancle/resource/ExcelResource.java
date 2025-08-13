package com.ewancle.resource;

import io.smallrye.mutiny.Multi;
import io.smallrye.mutiny.Uni;
import io.smallrye.mutiny.infrastructure.Infrastructure;

import io.smallrye.mutiny.unchecked.Unchecked;
import io.vertx.ext.web.FileUpload;
import io.vertx.mutiny.core.Vertx;
import io.vertx.mutiny.core.file.FileSystem;
import io.vertx.mutiny.core.buffer.Buffer;
import io.vertx.core.file.OpenOptions;

import jakarta.inject.Inject;
import jakarta.ws.rs.*;
import jakarta.ws.rs.Path;
import jakarta.ws.rs.core.MediaType;
import jakarta.ws.rs.core.HttpHeaders;
import org.jboss.resteasy.reactive.RestForm;
import org.jboss.resteasy.reactive.RestResponse;
import org.jboss.resteasy.reactive.RestResponse.ResponseBuilder;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

@Path("/excel")
public class ExcelResource {

    @Inject
    Vertx vertx;

    private final java.nio.file.Path uploadDir = java.nio.file.Path.of("uploads");

    public ExcelResource() throws IOException {
        if (!Files.exists(uploadDir)) {
            Files.createDirectories(uploadDir);
        }
    }

    private final String uploadsDir = System.getProperty("app.uploads.dir", "uploads");

    // 可选：自定义 worker 池（也可以使用 Infrastructure.getDefaultExecutor()）
    private final ExecutorService bgPool = Executors.newCachedThreadPool();

    public static class RowDto {
        public List<String> cells;
        public RowDto() {}
        public RowDto(List<String> cells) { this.cells = cells; }
    }

    /**
     * 1) 上传并解析 Excel（返回一个 Multi，每一项为一行）
     *    - Quarkus 已经把上传的 part 写入临时文件（传入为 java.io.File）
     *    - 解析工作在后台 worker 池执行，解析出的每一行 emit 出去（流式）
     */
    @POST
    @Path("/upload-parse")
    @Consumes(MediaType.MULTIPART_FORM_DATA)
    @Produces(MediaType.APPLICATION_JSON)
    public Multi<RowDto> uploadAndParse(@RestForm("file") File uploadedTempFile) {
        // 注意：不要在事件循环线程里执行 POI 操作
        return Multi.createFrom().emitter(emitter -> {
            bgPool.submit(() -> {
                try (InputStream is = new FileInputStream(uploadedTempFile)) {
                    // 判断是否 xlsx/xls：这里简单用文件名后缀，或用内容检测
                    String name = uploadedTempFile.getName().toLowerCase();
                    Workbook workbook;
                    if (name.endsWith(".xlsx") || name.endsWith(".xlsm")) {
                        // XSSF (xlsx)
                        workbook = new XSSFWorkbook(is);
                    } else if (name.endsWith(".xls")) {
                        workbook = WorkbookFactory.create(is);
                    } else {
                        // 尝试用 XSSFWorkbook 默认打开
                        workbook = WorkbookFactory.create(is);
                    }

                    // 以第一个 sheet 为例（或循环多个 sheet）
                    Sheet sheet = workbook.getNumberOfSheets() > 0 ? workbook.getSheetAt(0) : null;
                    if (sheet == null) {
                        emitter.complete();
                        workbook.close();
                        return;
                    }

                    for (Row row : sheet) {
                        // 解析一行为 List<String>（按 cell 类型转换）
                        List<String> cells = new ArrayList<>();
                        int maxCell = row.getLastCellNum();
                        for (int i = 0; i < maxCell; i++) {
                            Cell cell = row.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                            if (cell == null) {
                                cells.add(null);
                            } else {
                                switch (cell.getCellType()) {
                                    case STRING: cells.add(cell.getStringCellValue()); break;
                                    case NUMERIC:
                                        if (DateUtil.isCellDateFormatted(cell)) {
                                            cells.add(cell.getDateCellValue().toString());
                                        } else {
                                            cells.add(Double.toString(cell.getNumericCellValue()));
                                        }
                                        break;
                                    case BOOLEAN: cells.add(Boolean.toString(cell.getBooleanCellValue())); break;
                                    case FORMULA:
                                        // 获取公式结果（简化处理）
                                        try {
                                            cells.add(cell.getStringCellValue());
                                        } catch (Exception e) {
                                            cells.add(String.valueOf(cell.getNumericCellValue()));
                                        }
                                        break;
                                    case BLANK: cells.add(null); break;
                                    default: cells.add(cell.toString());
                                }
                            }
                        }
                        emitter.emit(new RowDto(cells));
                        // 这里可以考虑在大量数据时做一次性检查 emitter.isCancelled()
                        if (emitter.isCancelled()) {
                            break;
                        }
                    }

                    workbook.close();
                    emitter.complete();
                } catch (Throwable t) {
                    emitter.fail(t);
                } finally {
                    // 可选：删除 Quarkus 临时上传文件
                    try { Files.deleteIfExists(uploadedTempFile.toPath()); } catch (Exception ignore) {}
                }
            });
        });
    }

    /**
     * 2) 生成 Excel 并非阻塞流式下载
     *    - 生成工作在 worker 线程（将 Workbook 写入临时文件）
     *    - 然后用 Vert.x FileSystem 打开临时文件并通过 AsyncFile.toMulti() 流式返回
     */
    @GET
    @Path("/download-generated")
    public Uni<RestResponse<Multi<Buffer>>> generateAndDownload(@QueryParam("rows") @DefaultValue("1000") int rows) {
        // 生成临时文件路径
        String generatedName = UUID.randomUUID() + "-report.xlsx";
        FileSystem fs = vertx.fileSystem();

        // 1) 在 worker 线程生成 Excel 到临时文件（阻塞写）
        Uni<java.nio.file.Path> generateUni = Uni.createFrom().item(Unchecked.supplier(() -> {
            try (SXSSFWorkbook wb = new SXSSFWorkbook(100)) { // 使用流式 SXSSFWorkbook，防止内存爆炸
                Sheet sheet = wb.createSheet("sheet1");
                // 简单生成数据
                for (int r = 0; r < rows; r++) {
                    Row row = sheet.createRow(r);
                    for (int c = 0; c < 10; c++) {
                        Cell cell = row.createCell(c);
                        cell.setCellValue("R" + r + "C" + c);
                    }
                }
                //System.out.println("java.io.tmpdir = "+System.getProperty("java.io.tmpdir"));
                // 写入临时文件（阻塞）
                try (OutputStream os = Files.newOutputStream(
                        Paths.get(System.getProperty("java.io.tmpdir"), generatedName),
                        StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING)) {

                    wb.write(os);
                }
                // SXSSFWorkbook 推荐 dispose 临时资源
                wb.dispose();
                //wb.close();
                return  Paths.get(System.getProperty("java.io.tmpdir"), generatedName);
            } catch (Throwable ex) {
                // 确保删除临时文件（若已创建）
                try { Files.deleteIfExists(
                        Paths.get(System.getProperty("java.io.tmpdir"),
                                generatedName)); } catch (Exception ignore) {}
                throw new RuntimeException(ex);
            }
        })).runSubscriptionOn(Infrastructure.getDefaultExecutor());

        // 2) 生成后打开 AsyncFile 返回 Multi<Buffer>
        return generateUni.flatMap(path ->
                fs.open(path.toString(), new OpenOptions().setRead(true))
                        .map(asyncFile -> {
                            Multi<Buffer> body = asyncFile.toMulti()
                                    .onTermination().call(() -> {
                                        // 关闭句柄后删除临时文件（异步）
                                        return asyncFile.close().call(() -> {
                                            try { Files.deleteIfExists(path); }
                                            catch (Exception ignore) {}
                                            return Uni.createFrom().voidItem();
                                        });
                                    });

                            String attachmentName = "report.xlsx";
                            String mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                            return ResponseBuilder.ok(body)
                                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + attachmentName + "\"")
                                    .type(mime)
                                    .build();
                        })
        );
    }


    /**
     * 生成并下载 Excel（内存缓冲优化版）
     */
    @GET
    @Path("/download-generated-plus")
    public Uni<RestResponse<Multi<Buffer>>> downloadGenerated(@QueryParam("rows") @DefaultValue("1000") int rows) {
        // 1) Worker 线程生成 Excel 并返回 byte[]
        Uni<byte[]> excelBytes = Uni.createFrom().item(Unchecked.supplier(() -> {
            try (SXSSFWorkbook wb = new SXSSFWorkbook(100)) {
                Sheet sheet = wb.createSheet("sheet1");
                for (int r = 0; r < rows; r++) {
                    Row row = sheet.createRow(r);
                    for (int c = 0; c < 10; c++) {
                        Cell cell = row.createCell(c);
                        cell.setCellValue("R" + r + "C" + c);
                    }
                }
                try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
                    wb.write(baos);
                    wb.dispose();
                    return baos.toByteArray();
                }
            } catch (IOException e) {
                throw new RuntimeException("Excel generation failed", e);
            }
        })).runSubscriptionOn(Infrastructure.getDefaultExecutor());

        // 2) 分块流式返回
        return excelBytes.map(bytes -> {
            final int chunkSize = 64 * 1024; // 64KB per chunk

            Multi<Buffer> body = Multi.createFrom().emitter(emitter -> {
                try {
                    int offset = 0;
                    while (offset < bytes.length) {
                        int len = Math.min(chunkSize, bytes.length - offset);
                        //Buffer buf = Buffer.buffer(bytes, offset, len);
                        Buffer buf = Buffer.buffer(len).appendBytes(bytes, offset, len);
                        // 备选：Buffer buf = Buffer.buffer(Arrays.copyOfRange(bytes, offset, offset + len));
                        emitter.emit(buf);
                        offset += len;
                        if (emitter.isCancelled()) break;
                    }
                    emitter.complete();
                } catch (Throwable t) {
                    emitter.fail(t);
                }
            });

            String attachmentName = "report-" + UUID.randomUUID() + ".xlsx";
            String mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            return RestResponse.ResponseBuilder.ok(body)
                    .header(HttpHeaders.CONTENT_DISPOSITION,
                            "attachment; filename=\"" + attachmentName + "\"")
                    .type(mime)
                    .build();
        });
    }
}
