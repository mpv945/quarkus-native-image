package com.ewancle.resource;

import com.ewancle.model.UploadResponse;
import io.smallrye.mutiny.Multi;
import io.smallrye.mutiny.Uni;
import io.smallrye.mutiny.infrastructure.Infrastructure;
import io.vertx.core.file.OpenOptions;
import io.vertx.mutiny.core.Vertx;
import io.vertx.mutiny.core.buffer.Buffer;
import io.vertx.mutiny.core.file.AsyncFile;
import io.vertx.mutiny.core.file.FileSystem;
import jakarta.enterprise.context.ApplicationScoped;
import jakarta.inject.Inject;
import jakarta.ws.rs.*;
import jakarta.ws.rs.core.HttpHeaders;
import jakarta.ws.rs.core.MediaType;
import org.eclipse.microprofile.config.inject.ConfigProperty;
import org.jboss.resteasy.reactive.RestForm;
import org.jboss.resteasy.reactive.RestResponse;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.UUID;

@ApplicationScoped
//@RunOnVirtualThread 和响应式二选一
@Path("/files")
public class FileResource {

    // 上传：
    // curl -v -F "file=@/path/to/local-file.pdf" -F "filename=local-file.pdf" http://localhost:8080/files/upload
    //# 返回 JSON: { "originalName":"local-file.pdf", "storedName":"<uuid>-local-file.pdf", "size":12345 }

    // 下载：
    // curl -v -OJ http://localhost:8080/files/download/<storedName>  （注意 storedName 就是上传接口返回的 storedName）

    @Inject
    Vertx vertx;

    @ConfigProperty(name = "app.uploads.dir", defaultValue = "uploads")
    String uploadsDir;

    /**
     * 非阻塞上传：
     * 接收 multipart 的 file（这里用 java.io.File，Quarkus 会把 multipart 内容临时写到一个文件）
     * 建议同时传一个 "filename" 表单字段来保存原始文件名（客户端通常会这样做）。
     */
    @POST
    @Path("/upload")
    @Consumes(MediaType.MULTIPART_FORM_DATA)
    @Produces(MediaType.APPLICATION_JSON)
    public Uni<UploadResponse> upload(@RestForm("file") File uploadedTempFile,
                                      @RestForm("filename") String filenameFromForm) {
        final String originalName = (filenameFromForm != null && !filenameFromForm.isBlank())
                ? filenameFromForm
                : uploadedTempFile.getName();

        // 防止路径穿越
        final String safeOriginal = Paths.get(originalName).getFileName().toString();

        final String storedName = UUID.randomUUID() + "-" + safeOriginal;
        final String targetPath = uploadsDir + "/" + storedName;

        FileSystem fs = vertx.fileSystem();

        // 确保目录存在 -> 异步 copy -> 删除上传时的临时文件 -> 返回文件元信息
        return fs.mkdirs(uploadsDir)
                .chain(() -> fs.copy(uploadedTempFile.getAbsolutePath(), targetPath))
                // 删除 Quarkus 临时文件（非阻塞）
                .call(() -> fs.unlink(uploadedTempFile.getAbsolutePath()))
                .chain(() -> fs.props(targetPath))
                .map(props -> new UploadResponse(safeOriginal, storedName, props.size()));
    }

    /**
     * 非阻塞下载：
     * 传入之前返回的 storedName（UUID-原名）
     * 返回 Multi<Buffer>，Quarkus 将以流方式把内容写回客户端（application/octet-stream）
     */
    @GET
    @Path("/download/{storedName}")
    @Produces(MediaType.APPLICATION_OCTET_STREAM)
    public Multi<Buffer> download(@PathParam("storedName") String storedName) {
        final String safe = Paths.get(storedName).getFileName().toString();
        final String path = uploadsDir + "/" + safe;

        FileSystem fs = vertx.fileSystem();

        // open() -> Uni<AsyncFile>, transformToMulti(asyncFile -> asyncFile.toMulti())
        // 并在流终止时关闭文件句柄（防资源泄露）
        return fs.open(path, new OpenOptions().setRead(true))
                .onItem().transformToMulti(asyncFile ->
                        asyncFile.toMulti()
                                // 当 Multi 终止（完成/失败/取消）时，关闭文件（非阻塞 Uni）
                                .onTermination().call(asyncFile::close)
                );
    }


    @GET
    @Path("/download/{storedName}")
    public Uni<RestResponse<Multi<Buffer>>> download1(@PathParam("storedName") String storedName) {
        final String safe = Paths.get(storedName).getFileName().toString();
        final String path = uploadsDir + "/" + safe;

        // 从 storedName 截取原始文件名（假设格式 UUID-原名）
        String originalName = safe.contains("-") ? safe.substring(safe.indexOf('-') + 1) : safe;

        FileSystem fs = vertx.fileSystem();

        // 异步获取 MIME 类型（阻塞操作走工作线程）
        Uni<String> mimeUni = Uni.createFrom().item(() -> {
            try {
                String mime = Files.probeContentType(Paths.get(path));
                return mime != null ? mime : MediaType.APPLICATION_OCTET_STREAM;
            } catch (Exception e) {
                return MediaType.APPLICATION_OCTET_STREAM;
            }
        }).runSubscriptionOn(Infrastructure.getDefaultExecutor());

        return mimeUni
                .flatMap(mime ->
                        fs.open(path, new OpenOptions().setRead(true))
                                .map((AsyncFile asyncFile) -> {
                                    Multi<Buffer> body = asyncFile.toMulti()
                                            .onTermination().call(asyncFile::close);

                                    return RestResponse.ResponseBuilder
                                            .ok(body)
                                            .header(HttpHeaders.CONTENT_DISPOSITION,
                                                    "attachment; filename=\"" + originalName + "\"")
                                            .type(mime)
                                            .build();
                                })
                );
    }
}