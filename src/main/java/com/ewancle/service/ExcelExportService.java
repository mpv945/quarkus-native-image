package com.ewancle.service;

import com.ewancle.model.Person;
import io.smallrye.mutiny.Uni;
import io.smallrye.mutiny.unchecked.Unchecked;
import io.vertx.core.buffer.Buffer;
import jakarta.enterprise.context.ApplicationScoped;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;

@ApplicationScoped
public class ExcelExportService {

    public Uni<List<Person>> loadData(int rows) {
        return Uni.createFrom().item(() -> {
            List<Person> list = new ArrayList<>();
            IntStream.rangeClosed(1, rows).forEach(i -> list.add(
                    new Person((long) i, "Name" + i, "user" + i + "@example.com", 20 + (i % 30), LocalDateTime.now())
            ));
            return list;
        });
    }

    public Uni<Buffer> generateReactive(List<Person> people) {
        return Uni.createFrom().item(Unchecked.supplier(() -> {
            try (XSSFWorkbook workbook = new XSSFWorkbook();
                 ByteArrayOutputStream out = new ByteArrayOutputStream()
            ) {
                Sheet sheet = workbook.createSheet("People");
                createHeader(sheet);
                int rowIdx = 1;
                for (Person p : people) {
                    Row row = sheet.createRow(rowIdx++);
                    row.createCell(0).setCellValue(p.id());
                    row.createCell(1).setCellValue(p.name());
                    row.createCell(2).setCellValue(p.email());
                    row.createCell(3).setCellValue(p.age());
                    row.createCell(4).setCellValue(p.createdAt().toString());
                }
                workbook.write(out);
                return Buffer.buffer(out.toByteArray());
            } catch (Exception e) {
                throw new RuntimeException("Failed to generate Excel", e);
            }
        }));
    }

    private void createHeader(Sheet sheet) {
        Row header = sheet.createRow(0);
        String[] titles = {"ID", "Name", "Email", "Age", "Created At"};
        for (int i = 0; i < titles.length; i++) {
            header.createCell(i).setCellValue(titles[i]);
        }
    }
}
