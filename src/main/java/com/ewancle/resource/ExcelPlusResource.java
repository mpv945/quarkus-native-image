package com.ewancle.resource;

import com.ewancle.model.Employee;
import com.ewancle.service.ExcelService;
import io.smallrye.mutiny.Multi;
import io.smallrye.mutiny.Uni;
import io.vertx.core.buffer.Buffer;

import jakarta.inject.Inject;
import jakarta.ws.rs.*;
import jakarta.ws.rs.core.MediaType;
import jakarta.ws.rs.core.Response;
import jakarta.ws.rs.core.StreamingOutput;
import java.io.OutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

@Path("/excel1")
@Produces(MediaType.APPLICATION_JSON)
public class ExcelPlusResource {

    @Inject
    ExcelService excelService;

    // 方式1: 返回Multi<Buffer>的响应式流式下载
    @GET
    @Path("/stream-buffer")
    @Produces("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    public Multi<Buffer> downloadExcelStreamAsBuffer() {
        return excelService.generateExcelStreamAsBuffer()
                .onFailure().transform(throwable ->
                        new WebApplicationException("Excel生成失败: " + throwable.getMessage(), 500));
    }

    // 方式2: 返回Multi<byte[]>的响应式流式下载（兼容原有方式）
    @GET
    @Path("/stream")
    @Produces("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    public Multi<byte[]> downloadExcelStream() {
        return excelService.generateExcelStream()
                .onFailure().transform(throwable ->
                        new WebApplicationException("Excel生成失败: " + throwable.getMessage(), 500));
    }

    // 方式3: 返回Buffer的完整文件下载
    @GET
    @Path("/download-buffer")
    @Produces("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    public Uni<Response> downloadExcelAsBuffer() {
        return excelService.generateCompleteExcelAsBuffer()
                .map(buffer -> {
                    String filename = "员工信息_" +
                            LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss")) +
                            ".xlsx";

                    return Response.ok(buffer.getBytes())
                            .header("Content-Disposition", "attachment; filename=\"" + filename + "\"")
                            .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                            .header("Content-Length", buffer.length())
                            .build();
                })
                .onFailure().transform(throwable ->
                        new WebApplicationException("Excel下载失败: " + throwable.getMessage(), 500));
    }

    // 方式4: 直接返回Buffer（不包装在Response中）
    @GET
    @Path("/raw-buffer")
    @Produces("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    public Uni<Buffer> downloadRawBuffer() {
        return excelService.generateCompleteExcelAsBuffer()
                .onFailure().transform(throwable ->
                        new WebApplicationException("Excel生成失败: " + throwable.getMessage(), 500));
    }

    // 方式5: 完整文件下载（byte[]方式 - 推荐用于生产环境）
    @GET
    @Path("/download")
    @Produces("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    public Uni<Response> downloadExcel() {
        return excelService.generateCompleteExcel()
                .map(excelBytes -> {
                    String filename = "员工信息_" +
                            LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss")) +
                            ".xlsx";

                    return Response.ok(excelBytes)
                            .header("Content-Disposition", "attachment; filename=\"" + filename + "\"")
                            .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                            .header("Content-Length", excelBytes.length)
                            .build();
                })
                .onFailure().transform(throwable ->
                        new WebApplicationException("Excel下载失败: " + throwable.getMessage(), 500));
    }

    // 方式6: 使用StreamingOutput配合Buffer进行流式输出
    @GET
    @Path("/streaming-buffer")
    @Produces("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    public Uni<Response> downloadExcelStreamingBuffer() {
        return Uni.createFrom().item(() -> {
            StreamingOutput streamingOutput = (OutputStream output) -> {
                excelService.generateExcelStreamAsBuffer()
                        .subscribe().with(
                                buffer -> {
                                    try {
                                        output.write(buffer.getBytes());
                                        output.flush();
                                    } catch (Exception e) {
                                        throw new RuntimeException(e);
                                    }
                                },
                                failure -> {
                                    throw new RuntimeException(failure);
                                }
                        );
            };

            String filename = "员工信息流式Buffer_" +
                    LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss")) +
                    ".xlsx";

            return Response.ok(streamingOutput)
                    .header("Content-Disposition", "attachment; filename=\"" + filename + "\"")
                    .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    .header("Transfer-Encoding", "chunked")
                    .build();
        });
    }

    // 获取数据预览
    @GET
    @Path("/preview")
    public Multi<Employee> getDataPreview() {
        return excelService.getEmployeeStream();
    }
}
