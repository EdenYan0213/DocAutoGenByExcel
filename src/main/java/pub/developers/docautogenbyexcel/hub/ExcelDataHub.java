package pub.developers.docautogenbyexcel.hub;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import pub.developers.docautogenbyexcel.model.ModuleData;
import pub.developers.docautogenbyexcel.model.Requirement;
import pub.developers.docautogenbyexcel.model.TestCase;
import pub.developers.docautogenbyexcel.model.TestResult;
import pub.developers.docautogenbyexcel.reader.ExcelReader;
import pub.developers.docautogenbyexcel.reader.TableDataReader;
import pub.developers.docautogenbyexcel.reader.TableDataReader.BasicInfoData;
import pub.developers.docautogenbyexcel.reader.TableDataReader.ListTableData;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Concrete DataHub implementation backed by Excel files.
 */
@Component
public class ExcelDataHub implements DataHub {

    private final ExcelReader excelReader;
    private final TableDataReader tableDataReader;

    public ExcelDataHub() {
        this.excelReader = new ExcelReader();
        this.tableDataReader = new TableDataReader();
    }

    @Override
    public Map<String, ModuleData> loadModuleData(String excelPath) throws Exception {
        return excelReader.readExcel(excelPath);
    }

    @Override
    public Map<String, BasicInfoData> loadBasicInfo(String excelPath) throws Exception {
        return tableDataReader.readBasicInfo(excelPath);
    }

    @Override
    public Map<String, ListTableData> loadListTables(String excelPath) throws Exception {
        return tableDataReader.readAllListTableData(excelPath);
    }

    @Override
    public List<Requirement> loadRequirements(String excelPath) throws Exception {
        List<Requirement> requirements = new ArrayList<>();
        try (Workbook workbook = openWorkbook(excelPath)) {
            Sheet sheet = findSheetByHeaders(workbook, Set.of("ReqID", "需求ID"));
            if (sheet == null) {
                return requirements;
            }

            Map<String, Integer> header = readHeaderMap(sheet);
            int reqIdCol = findColumn(header, "ReqID", "需求ID", "RequirementId");
            int reqTitleCol = findColumn(header, "ReqTitle", "需求标题", "需求名称");
            int descCol = findColumn(header, "Description", "需求描述");
            int priorityCol = findColumn(header, "Priority", "优先级");

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    continue;
                }
                String reqId = getCellValue(row.getCell(reqIdCol));
                if (reqId == null || reqId.isBlank()) {
                    continue;
                }
                Requirement req = new Requirement();
                req.setRequirementId(reqId.trim());
                req.setRequirementNumber(reqId.trim());
                req.setRequirementName(safeCell(row, reqTitleCol));
                req.setDescription(safeCell(row, descCol));
                String priority = safeCell(row, priorityCol);
                if ("高".equals(priority)) {
                    req.setPriority(Requirement.Priority.HIGH);
                } else if ("低".equals(priority)) {
                    req.setPriority(Requirement.Priority.LOW);
                } else {
                    req.setPriority(Requirement.Priority.MEDIUM);
                }
                requirements.add(req);
            }
        }
        return requirements;
    }

    @Override
    public List<TestCase> loadTestCases(String excelPath) throws Exception {
        List<TestCase> result = new ArrayList<>();
        Map<String, ModuleData> moduleDataMap = excelReader.readExcel(excelPath);
        for (ModuleData moduleData : moduleDataMap.values()) {
            result.addAll(moduleData.getTestCases());
        }
        return result;
    }

    @Override
    public List<TestResult> loadTestResults(String excelPath) throws Exception {
        List<TestResult> results = new ArrayList<>();
        try (Workbook workbook = openWorkbook(excelPath)) {
            Sheet sheet = findSheetByHeaders(workbook, Set.of("TCID", "测试用例标识", "标识"));
            if (sheet == null) {
                return results;
            }

            Map<String, Integer> header = readHeaderMap(sheet);
            int tcidCol = findColumn(header, "TCID", "测试用例标识", "标识");
            int execDateCol = findColumn(header, "ExecDate", "执行日期", "Date");
            int resultCol = findColumn(header, "Result", "测试结论", "结论");
            int defectCol = findColumn(header, "DefectID", "缺陷标识", "缺陷ID");

            if (resultCol < 0) {
                return results;
            }

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    continue;
                }
                String tcid = safeCell(row, tcidCol);
                String verdict = safeCell(row, resultCol);
                if (tcid.isBlank() || verdict.isBlank()) {
                    continue;
                }

                TestResult tr = new TestResult();
                tr.setTestCaseId(tcid);
                tr.setExecutionDate(safeCell(row, execDateCol));
                tr.setVerdict(verdict);
                tr.setDefectId(safeCell(row, defectCol));
                results.add(tr);
            }
        }
        return results;
    }

    @Override
    public Map<String, String> loadConfig(String excelPath) throws Exception {
        Map<String, String> config = new LinkedHashMap<>();
        try (Workbook workbook = openWorkbook(excelPath)) {
            Sheet meta = workbook.getSheet("Meta");
            if (meta == null) {
                meta = findSheetByHeaders(workbook, Set.of("Key", "键"));
            }
            if (meta == null) {
                return config;
            }

            Map<String, Integer> header = readHeaderMap(meta);
            int keyCol = findColumn(header, "Key", "键");
            int valueCol = findColumn(header, "Value", "值");
            if (keyCol < 0 || valueCol < 0) {
                return config;
            }

            for (int r = 1; r <= meta.getLastRowNum(); r++) {
                Row row = meta.getRow(r);
                if (row == null) {
                    continue;
                }
                String key = safeCell(row, keyCol);
                if (key.isBlank()) {
                    continue;
                }
                config.put(key, safeCell(row, valueCol));
            }
        }
        return config;
    }

    @Override
    public void writeRequirements(String excelPath, List<Requirement> requirements) throws Exception {
        Workbook workbook = null;
        try {
            workbook = openOrCreateWorkbook(excelPath);
            Sheet sheet = workbook.getSheet("Data_SRS_Requirements");
            if (sheet == null) {
                sheet = workbook.createSheet("Data_SRS_Requirements");
            }

            clearSheet(sheet);
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("ReqID");
            header.createCell(1).setCellValue("ReqTitle");
            header.createCell(2).setCellValue("Description");
            header.createCell(3).setCellValue("Priority");

            int rowNum = 1;
            for (Requirement req : requirements) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(nullToEmpty(req.getRequirementId()));
                row.createCell(1).setCellValue(nullToEmpty(req.getRequirementName()));
                row.createCell(2).setCellValue(nullToEmpty(req.getDescription()));
                row.createCell(3).setCellValue(req.getPriority() == null ? "中" : req.getPriority().getDescription());
            }

            persistWorkbook(workbook, excelPath);
        } finally {
            if (workbook != null) {
                workbook.close();
            }
        }
    }

    @Override
    public ImportSummary writeTestCases(String excelPath, List<TestCase> testCases) throws Exception {
        ImportSummary summary = new ImportSummary();
        List<Requirement> requirements = loadRequirements(excelPath);
        Set<String> reqIds = new HashSet<>();
        for (Requirement requirement : requirements) {
            reqIds.add(requirement.getRequirementId());
        }

        Workbook workbook = null;
        try {
            workbook = openOrCreateWorkbook(excelPath);
            Sheet sheet = workbook.getSheet("Data_STD_TestCases");
            if (sheet == null) {
                sheet = workbook.createSheet("Data_STD_TestCases");
                Row header = sheet.createRow(0);
                header.createCell(0).setCellValue("TCID");
                header.createCell(1).setCellValue("ReqID");
                header.createCell(2).setCellValue("Procedure");
                header.createCell(3).setCellValue("ExpectedResult");
            }

            int rowNum = sheet.getLastRowNum() + 1;
            for (TestCase tc : testCases) {
                String tcid = choose(tc.getColumnValue("TCID"), tc.getColumnValue("测试用例标识"), tc.getId());
                String reqId = choose(tc.getColumnValue("ReqID"), tc.getColumnValue("需求ID"), tc.getColumnValue("需求标识"));
                if (tcid.isBlank()) {
                    summary.addError("测试用例缺少TCID");
                    summary.setFailedCount(summary.getFailedCount() + 1);
                    continue;
                }
                if (!reqId.isBlank() && !reqIds.contains(reqId)) {
                    summary.addWarning("测试用例 " + tcid + " 关联了不存在的需求 " + reqId);
                }

                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(tcid);
                row.createCell(1).setCellValue(reqId);
                row.createCell(2).setCellValue(
                        choose(tc.getColumnValue("Procedure"), tc.getColumnValue("测试步骤"), tc.getContent()));
                row.createCell(3).setCellValue(
                        choose(tc.getColumnValue("ExpectedResult"), tc.getColumnValue("预期结果"), tc.getCriteria()));
                summary.setSuccessCount(summary.getSuccessCount() + 1);
            }

            persistWorkbook(workbook, excelPath);
            return summary;
        } finally {
            if (workbook != null) {
                workbook.close();
            }
        }
    }

    @Override
    public ImportSummary appendTestResults(String excelPath, List<TestResult> testResults) throws Exception {
        ImportSummary summary = new ImportSummary();
        List<TestCase> cases = loadTestCases(excelPath);
        Set<String> caseIds = new HashSet<>();
        for (TestCase tc : cases) {
            String tcid = choose(tc.getColumnValue("TCID"), tc.getColumnValue("测试用例标识"), tc.getId());
            if (!tcid.isBlank()) {
                caseIds.add(tcid);
            }
        }

        Workbook workbook = null;
        try {
            workbook = openOrCreateWorkbook(excelPath);
            Sheet sheet = workbook.getSheet("Data_STR_TestResults");
            if (sheet == null) {
                sheet = workbook.createSheet("Data_STR_TestResults");
                Row header = sheet.createRow(0);
                header.createCell(0).setCellValue("TCID");
                header.createCell(1).setCellValue("ExecDate");
                header.createCell(2).setCellValue("Result");
                header.createCell(3).setCellValue("DefectID");
            }

            int rowNum = sheet.getLastRowNum() + 1;
            for (TestResult tr : testResults) {
                String tcid = nullToEmpty(tr.getTestCaseId());
                if (tcid.isBlank()) {
                    summary.addError("测试结果缺少TCID");
                    summary.setFailedCount(summary.getFailedCount() + 1);
                    continue;
                }
                if (!caseIds.contains(tcid)) {
                    summary.addError("测试结果TCID不存在: " + tcid);
                    summary.setFailedCount(summary.getFailedCount() + 1);
                    continue;
                }

                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(tcid);
                row.createCell(1).setCellValue(nullToEmpty(tr.getExecutionDate()));
                row.createCell(2).setCellValue(nullToEmpty(tr.getVerdict()));
                row.createCell(3).setCellValue(nullToEmpty(tr.getDefectId()));
                summary.setSuccessCount(summary.getSuccessCount() + 1);
            }

            persistWorkbook(workbook, excelPath);
            return summary;
        } finally {
            if (workbook != null) {
                workbook.close();
            }
        }
    }

    private Workbook openWorkbook(String excelPath) throws IOException {
        try (FileInputStream fis = new FileInputStream(excelPath)) {
            return new XSSFWorkbook(fis);
        }
    }

    private Workbook openOrCreateWorkbook(String excelPath) throws IOException {
        File file = new File(excelPath);
        if (file.exists()) {
            try (FileInputStream fis = new FileInputStream(file)) {
                return new XSSFWorkbook(fis);
            }
        }
        return new XSSFWorkbook();
    }

    private void persistWorkbook(Workbook workbook, String excelPath) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(excelPath)) {
            workbook.write(fos);
        }
    }

    private Sheet findSheetByHeaders(Workbook workbook, Set<String> mustHaveOneOf) {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            Row header = sheet.getRow(0);
            if (header == null) {
                continue;
            }
            Set<String> names = new LinkedHashSet<>();
            for (int c = 0; c < header.getLastCellNum(); c++) {
                String value = getCellValue(header.getCell(c));
                if (!value.isBlank()) {
                    names.add(value);
                }
            }
            for (String candidate : mustHaveOneOf) {
                if (names.contains(candidate)) {
                    return sheet;
                }
            }
        }
        return null;
    }

    private Map<String, Integer> readHeaderMap(Sheet sheet) {
        Map<String, Integer> map = new HashMap<>();
        Row row = sheet.getRow(0);
        if (row == null) {
            return map;
        }
        for (int c = 0; c < row.getLastCellNum(); c++) {
            String name = getCellValue(row.getCell(c));
            if (!name.isBlank()) {
                map.put(name, c);
            }
        }
        return map;
    }

    private int findColumn(Map<String, Integer> header, String... names) {
        for (String name : names) {
            Integer idx = header.get(name);
            if (idx != null) {
                return idx;
            }
        }
        return -1;
    }

    private String safeCell(Row row, int index) {
        if (index < 0 || row == null) {
            return "";
        }
        return nullToEmpty(getCellValue(row.getCell(index))).trim();
    }

    private String choose(String... candidates) {
        for (String value : candidates) {
            if (value != null && !value.isBlank()) {
                return value.trim();
            }
        }
        return "";
    }

    private String nullToEmpty(String s) {
        return s == null ? "" : s;
    }

    private String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        CellType type = cell.getCellType();
        if (type == CellType.STRING) {
            return cell.getStringCellValue();
        }
        if (type == CellType.NUMERIC) {
            double v = cell.getNumericCellValue();
            if (v == Math.floor(v)) {
                return String.valueOf((long) v);
            }
            return String.valueOf(v);
        }
        if (type == CellType.BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        }
        if (type == CellType.FORMULA) {
            try {
                return cell.getStringCellValue();
            } catch (Exception ignored) {
                try {
                    double v = cell.getNumericCellValue();
                    if (v == Math.floor(v)) {
                        return String.valueOf((long) v);
                    }
                    return String.valueOf(v);
                } catch (Exception ex) {
                    return "";
                }
            }
        }
        return "";
    }

    private void clearSheet(Sheet sheet) {
        int last = sheet.getLastRowNum();
        for (int i = last; i >= 0; i--) {
            Row row = sheet.getRow(i);
            if (row != null) {
                sheet.removeRow(row);
            }
        }
    }
}
