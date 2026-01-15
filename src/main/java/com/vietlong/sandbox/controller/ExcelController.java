package com.vietlong.sandbox.controller;

import com.vietlong.sandbox.service.ExcelParserService;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.Parameter;
import io.swagger.v3.oas.annotations.media.Content;
import io.swagger.v3.oas.annotations.tags.Tag;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;
import java.util.Map;

@Tag(name = "Excel Parser", description = "API xử lý file Excel và chuyển đổi sang JSON động")
@RestController
@RequestMapping("/api/excel")
public class ExcelController {

    @Autowired
    private ExcelParserService excelParserService;

    @Operation(
            summary = "Upload và parse file Excel"
    )
    @PostMapping(value = "/parse-to-json", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<List<Map<String, Object>>> uploadAndParseExcelToJson(
            @Parameter(
                    description = "File Excel cần parse (.xlsx)",
                    content = @Content(mediaType = MediaType.MULTIPART_FORM_DATA_VALUE)
            )
            @RequestPart("file") MultipartFile file,

            @Parameter(
                    description = "Danh sách index các cột cần lấy (bắt đầu từ 0). Nếu để trống sẽ lấy tất cả.",
                    required = false
            )
            @RequestParam(value = "columnIndexes", required = false) List<Integer> columnIndexes
    ) {
        if (file.isEmpty() || file.getOriginalFilename() == null || !file.getOriginalFilename().endsWith(".xlsx")) {
            return ResponseEntity.badRequest().build();
        }

        if (columnIndexes != null && !columnIndexes.isEmpty()) {
            return ResponseEntity.ok(excelParserService.parseToJson(file, columnIndexes));
        }

        return ResponseEntity.ok(excelParserService.parseToJson(file));
    }
}