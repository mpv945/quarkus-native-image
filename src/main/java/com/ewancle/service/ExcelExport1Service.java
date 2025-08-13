package com.ewancle.service;

import io.smallrye.mutiny.Uni;
import io.vertx.core.Vertx;
import io.vertx.core.file.FileSystem;
import io.vertx.core.file.OpenOptions;
import jakarta.enterprise.context.ApplicationScoped;
import jakarta.inject.Inject;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;

@ApplicationScoped
public class ExcelExport1Service {

    @Inject
    Vertx vertx;

    /**
     * 上传 → 解析 → 生成新 Excel（全 Reactive）
     */
    public Uni<Path> processExcelReactive(InputStream uploadedFile) {
        return Uni.createFrom().<Path>emitter(em -> {
            vertx.<Path>executeBlocking(
                    promise -> {
                        try (// 解析 Excel
                             XSSFWorkbook workbook = new XSSFWorkbook(uploadedFile);){


                            // 在这里处理解析结果，比如新增一行
                            var sheet = workbook.getSheetAt(0);
                            var row = sheet.createRow(sheet.getLastRowNum() + 1);
                            row.createCell(0).setCellValue("新增数据");

                            // 写入新 Excel
                            Path outputFile = Files.createTempFile("processed-", ".xlsx");
                            try (OutputStream os = Files.newOutputStream(outputFile)) {
                                workbook.write(os);
                            }
                            promise.complete(outputFile);
                        } catch (Exception e) {
                            promise.fail(e);
                        }
                    },
                    false, // false = 使用共享 worker pool
                    ar -> {
                        if (ar.succeeded()) {
                            em.complete(ar.result());
                        } else {
                            em.fail(ar.cause());
                        }
                    }
            );
        });
    }
}