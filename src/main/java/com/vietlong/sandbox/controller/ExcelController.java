package com.vietlong.sandbox.controller;

import com.vietlong.sandbox.service.ExcelParserService;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.Parameter;
import io.swagger.v3.oas.annotations.media.Content;
import io.swagger.v3.oas.annotations.media.ExampleObject;
import io.swagger.v3.oas.annotations.media.Schema;
import io.swagger.v3.oas.annotations.responses.ApiResponse;
import io.swagger.v3.oas.annotations.responses.ApiResponses;
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
        summary = "Upload và parse file Excel",
        description = """
            Upload file Excel (.xlsx) và tự động chuyển đổi sang JSON.
            
            **Cách hoạt động:**
            - Dòng đầu tiên (header) sẽ được dùng làm key cho JSON
            - Header sẽ được chuẩn hóa: "Mã Sản Phẩm" → "MA_SAN_PHAM"
            - Các dòng tiếp theo sẽ được parse thành JSON objects
            - Hỗ trợ tự động nhận diện kiểu dữ liệu (text, số, ngày, boolean)
            
            **Lưu ý:**
            - Chỉ hỗ trợ file .xlsx (Excel 2007+)
            - Chỉ đọc sheet đầu tiên
            - Dòng trống sẽ tự động bị bỏ qua
            """
    )
    @PostMapping(value = "/parse-to-json", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<List<Map<String, Object>>> uploadAndParseExcelToJson(
        @Parameter(
            content = @Content(mediaType = MediaType.MULTIPART_FORM_DATA_VALUE)
        )
        @RequestParam("file") MultipartFile file,
        @RequestParam("columnIndexes") List<Integer> columnIndexes
    ) {
        if (file.isEmpty() || !file.getOriginalFilename().endsWith(".xlsx")) {
            return ResponseEntity.badRequest().body(null);
        }

        if (!columnIndexes.isEmpty()){
            return ResponseEntity.ok(excelParserService.parseToJson(file, columnIndexes));
        }

        return ResponseEntity.ok(excelParserService.parseToJson(file));
    }
}