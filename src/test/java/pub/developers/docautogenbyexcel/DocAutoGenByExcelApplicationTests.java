package pub.developers.docautogenbyexcel;

import org.junit.jupiter.api.Test;
import org.springframework.boot.test.mock.mockito.MockBean;
import pub.developers.docautogenbyexcel.repository.DocumentRepository;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.boot.autoconfigure.EnableAutoConfiguration;
import org.springframework.boot.autoconfigure.jdbc.DataSourceAutoConfiguration;

@SpringBootTest
@EnableAutoConfiguration(exclude = {DataSourceAutoConfiguration.class})
class DocAutoGenByExcelApplicationTests {

	@MockBean
	private DocumentRepository documentRepository;

	@Test
	void contextLoads() {
	}

}
