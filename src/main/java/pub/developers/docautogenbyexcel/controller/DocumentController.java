package pub.developers.docautogenbyexcel.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import pub.developers.docautogenbyexcel.service.DocumentService;
import pub.developers.docautogenbyexcel.service.DocumentService.DocumentInfo;
import pub.developers.docautogenbyexcel.service.DocumentService.ProcessResult;

import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Map;

/**
 * 文档处理 REST API
 */
@RestController
@RequestMapping("/api/documents")
@CrossOrigin(origins = "*")
public class DocumentController {

    @Autowired
    private DocumentService documentService;

    /**
     * 上传Excel和Word文件进行处理
     * 
     * POST /api/documents/process
     * Content-Type: multipart/form-data
     * 
     * @param excelFile Excel数据文件
     * @param wordFile Word模板文件
     */
    @PostMapping("/process")
    public ResponseEntity<?> processDocuments(
            @RequestParam("excel") MultipartFile excelFile,
            @RequestParam("word") MultipartFile wordFile) {
        
        // 验证文件
        if (excelFile.isEmpty() || wordFile.isEmpty()) {
            return ResponseEntity.badRequest()
                .body(Map.of("error", "请上传Excel和Word文件"));
        }
        
        String excelName = excelFile.getOriginalFilename();
        String wordName = wordFile.getOriginalFilename();
        
        if (excelName == null || !excelName.matches(".*\\.xlsx?$")) {
            return ResponseEntity.badRequest()
                .body(Map.of("error", "请上传有效的Excel文件(.xlsx或.xls)"));
        }
        
        if (wordName == null || !wordName.matches(".*\\.docx?$")) {
            return ResponseEntity.badRequest()
                .body(Map.of("error", "请上传有效的Word文件(.docx或.doc)"));
        }
        
        try {
            ProcessResult result = documentService.processDocuments(
                excelFile.getInputStream(), excelName,
                wordFile.getInputStream(), wordName
            );
            
            return ResponseEntity.ok(Map.of(
                "success", true,
                "message", result.message(),
                "outputId", result.outputId(),
                "outputFileName", result.outputFileName(),
                "moduleCount", result.moduleCount(),
                "downloadUrl", "/api/documents/download/" + result.outputFileName()
            ));
            
        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                .body(Map.of(
                    "error", "处理失败: " + e.getMessage(),
                    "success", false
                ));
        }
    }

    /**
     * 下载处理后的文档
     * 
     * GET /api/documents/download/{fileName}
     */
    @GetMapping("/download/{fileName}")
    public ResponseEntity<?> downloadDocument(@PathVariable String fileName) {
        try {
            byte[] content = documentService.getOutputDocument(fileName);
            
            String encodedFileName = URLEncoder.encode(fileName, StandardCharsets.UTF_8)
                .replaceAll("\\+", "%20");
            
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            headers.setContentDisposition(ContentDisposition.attachment()
                .filename(encodedFileName, StandardCharsets.UTF_8)
                .build());
            headers.setContentLength(content.length);
            
            return new ResponseEntity<>(content, headers, HttpStatus.OK);
            
        } catch (Exception e) {
            return ResponseEntity.notFound().build();
        }
    }

    /**
     * 预览文档（返回文件内容用于在线预览）
     * 
     * GET /api/documents/preview/{fileName}
     */
    @GetMapping("/preview/{fileName}")
    public ResponseEntity<?> previewDocument(@PathVariable String fileName) {
        try {
            byte[] content = documentService.getOutputDocument(fileName);
            
            HttpHeaders headers = new HttpHeaders();
            // 使用适合预览的Content-Type
            headers.setContentType(MediaType.parseMediaType(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
            headers.setContentLength(content.length);
            // 设置为inline以便浏览器预览
            headers.setContentDisposition(ContentDisposition.inline()
                .filename(fileName, StandardCharsets.UTF_8)
                .build());
            
            return new ResponseEntity<>(content, headers, HttpStatus.OK);
            
        } catch (Exception e) {
            return ResponseEntity.notFound().build();
        }
    }

    /**
     * 获取已处理的文档列表
     * 
     * GET /api/documents/list
     */
    @GetMapping("/list")
    public ResponseEntity<?> listDocuments() {
        try {
            List<DocumentInfo> documents = documentService.listOutputDocuments();
            return ResponseEntity.ok(Map.of(
                "success", true,
                "documents", documents,
                "total", documents.size()
            ));
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                .body(Map.of("error", e.getMessage()));
        }
    }

    /**
     * 删除文档
     * 
     * DELETE /api/documents/{fileName}
     */
    @DeleteMapping("/{fileName}")
    public ResponseEntity<?> deleteDocument(@PathVariable String fileName) {
        boolean deleted = documentService.deleteDocument(fileName);
        if (deleted) {
            return ResponseEntity.ok(Map.of(
                "success", true,
                "message", "文档已删除"
            ));
        } else {
            return ResponseEntity.notFound().build();
        }
    }

    /**
     * 清理旧文件
     * 
     * POST /api/documents/cleanup
     */
    @PostMapping("/cleanup")
    public ResponseEntity<?> cleanupOldFiles(
            @RequestParam(defaultValue = "7") int daysToKeep) {
        int deletedCount = documentService.cleanupOldFiles(daysToKeep);
        return ResponseEntity.ok(Map.of(
            "success", true,
            "deletedCount", deletedCount,
            "message", "已清理 " + deletedCount + " 个旧文件"
        ));
    }
}

