package pub.developers.docautogenbyexcel;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

/**
 * Spring Boot应用入口（可选）
 * 主要入口为ExcelToWordTool类
 */
@SpringBootApplication
public class DocAutoGenByExcelApplication {

    public static void main(String[] args) {
        // 如果作为Spring Boot应用运行，使用SpringApplication
        // 否则直接使用ExcelToWordTool.main(args)
        if (args.length > 0 && !args[0].startsWith("--spring")) {
            ExcelToWordTool.main(args);
        } else {
            SpringApplication.run(DocAutoGenByExcelApplication.class, args);
        }
    }

}
