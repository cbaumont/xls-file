package local.xls_file;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App {
	public static void main(String[] args) {
		// Criando workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		// Criando spreadSheet
		XSSFSheet spreadSheet = workbook.createSheet("Planilha Teste");
		// Map com os valores a serem inseridos
		Map<String, List<? extends Object>> dados = new LinkedHashMap<>();
		List<String> header = Arrays.asList("Valor 01", "Valor 02", "Valor 03");
		List<String> firstLine = Arrays.asList("Aluguel", "100,00", "20,00");
		dados.put("1", header);
		dados.put("2", firstLine);
		// Set com as linhas
		Set<String> keySet = dados.keySet();
		int rowid = 0;
		// Iterando para inserir os valores
		for (String indice : keySet) {
			XSSFRow row = spreadSheet.createRow(rowid++);
			@SuppressWarnings("unchecked")
			List<String> lines = (List<String>) dados.get(indice);

			int cellId = 0;

			for (String line : lines) {
				Cell cell = row.createCell(cellId++);
				cell.setCellValue(line);
			}
		}
		// Escrevendo o workbook no novo arquivo
		try {
			FileOutputStream planilha = new FileOutputStream(new File("PlanilhaExemplo.xlsx"));
			workbook.write(planilha);
			planilha.close();
			workbook.close();
			System.out.println("Sucesso na criacao do arquivo!");
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
}
