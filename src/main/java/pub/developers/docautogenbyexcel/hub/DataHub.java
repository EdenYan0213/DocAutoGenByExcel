package pub.developers.docautogenbyexcel.hub;

import pub.developers.docautogenbyexcel.model.ModuleData;
import pub.developers.docautogenbyexcel.model.Requirement;
import pub.developers.docautogenbyexcel.model.TestCase;
import pub.developers.docautogenbyexcel.model.TestResult;
import pub.developers.docautogenbyexcel.reader.TableDataReader.BasicInfoData;
import pub.developers.docautogenbyexcel.reader.TableDataReader.ListTableData;

import java.util.List;
import java.util.Map;

/**
 * Excel data hub abstraction.
 * Provides a unified data view for readers and generators.
 */
public interface DataHub {

    Map<String, ModuleData> loadModuleData(String excelPath) throws Exception;

    Map<String, BasicInfoData> loadBasicInfo(String excelPath) throws Exception;

    Map<String, ListTableData> loadListTables(String excelPath) throws Exception;

    List<Requirement> loadRequirements(String excelPath) throws Exception;

    List<TestCase> loadTestCases(String excelPath) throws Exception;

    List<TestResult> loadTestResults(String excelPath) throws Exception;

    Map<String, String> loadConfig(String excelPath) throws Exception;

    void writeRequirements(String excelPath, List<Requirement> requirements) throws Exception;

    ImportSummary writeTestCases(String excelPath, List<TestCase> testCases) throws Exception;

    ImportSummary appendTestResults(String excelPath, List<TestResult> testResults) throws Exception;
}
