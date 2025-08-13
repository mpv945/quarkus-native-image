package com.ewancle.resource;

import com.ewancle.service.ExcelExport1Service;
import com.ewancle.service.ExcelExportService;
import io.smallrye.mutiny.Uni;
import io.smallrye.mutiny.unchecked.Unchecked;
import jakarta.inject.Inject;
import jakarta.ws.rs.*;
import jakarta.ws.rs.core.Context;
import jakarta.ws.rs.core.MediaType;
import jakarta.ws.rs.core.Response;
import org.jboss.resteasy.reactive.RestForm;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

@Path("/export")
public class ExceChatgptlResource {

    @Inject ExcelExportService service;
    @Inject ExcelExport1Service excelService;

    @GET
    @Path("/excel")
    @Produces("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    public Uni<Response> exportExcel(@QueryParam("rows") @DefaultValue("10000") int rows) {
        String filename = "people-" + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss")) + ".xlsx";
        return service.loadData(rows)
                .flatMap(list -> service.generateReactive(list)
                        .map(fileBuffer -> Response.ok(fileBuffer)
                                .header("Content-Disposition", "attachment; filename=\"" + filename + "\"")
                                .type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                                .build()
                        )
                );
    }

    /*@POST
    @Path("/upload")
    @Consumes(MediaType.MULTIPART_FORM_DATA)
    @Produces(MediaType.APPLICATION_OCTET_STREAM)
    public Uni<Response> uploadAndProcess(@Context RoutingContext rc) {
        // 获取上传的文件（Reactive 方式）
        List<FileUpload> uploads = rc.fileUploads();
        if (uploads.isEmpty()) {
            return Uni.createFrom().item(Response.status(400).entity("No file uploaded").build());
        }

        FileUpload upload = uploads.getFirst();
        java.nio.file.Path uploadedPath = Paths.get(upload.uploadedFileName());

        // Reactive 调用解析逻辑
        return excelService.processExcelReactive(upload.f)
                .map(newFilePath -> {
                    return Response.ok()
                            .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=processed.xlsx")
                            .entity((StreamingOutput) output -> Files.copy(newFilePath, output))
                            .build();
                });
    }*/

    @POST
    @Path("/upload")
    @Consumes(MediaType.MULTIPART_FORM_DATA)
    public Uni<String> uploadExcel(@Context io.vertx.mutiny.ext.web.RoutingContext rc) {
        return Uni.createFrom().item(Unchecked.supplier(() -> {
            var uploads = rc.fileUploads();
            if (uploads.isEmpty()) throw new RuntimeException("No file uploaded");
            var upload = uploads.getFirst();
            return upload.uploadedFileName(); // 临时路径
        })).flatMap(Unchecked.function(pathStr -> {
            java.nio.file.Path path = java.nio.file.Path.of(pathStr);
            try (InputStream in = Files.newInputStream(path)) {
                return excelService.processExcelReactive(in);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        })).map(p -> "File processed: " + p.toAbsolutePath());
    }

    @POST
    @Path("/upload1")
    @Consumes(MediaType.MULTIPART_FORM_DATA)
    public Uni<String> uploadExcel1(@RestForm("file") InputStream uploadedFile) {
        return excelService.processExcelReactive(uploadedFile)
                .map(path -> "File processed: " + path.toAbsolutePath());

    }

}
