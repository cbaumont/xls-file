package local.xls_file.utils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlsUtils {
	
	private static XSSFWorkbook workbook = new XSSFWorkbook();
	private static XSSFSheet spreadSheet = workbook.createSheet("Valores");
	private static Map<String, List<String>> valores = new LinkedHashMap<String, List<String>>();
	
	public static void geraCabecalho(Map<String, List<String>> cabecalhoValores) {
		System.out.println("Gerando cabecalho");
		if (spreadSheet.getPhysicalNumberOfRows() == 0 && !cabecalhoValores.isEmpty()) {
			geraLinhas(cabecalhoValores);
		}
	}

	public static void geraLinhas(Map<String, List<String>> linhasValores) {
		System.out.println("Gerando linhas");
		if (!linhasValores.isEmpty()) {
			valores.putAll(linhasValores);
		}
	}

	public static void valoresPlanilha() {
		System.out.println("Iterando para inserir os valores");
		Set<String> keySet = valores.keySet();
		int rowid = spreadSheet.getLastRowNum();
		for (String indice : keySet) {
			XSSFRow row = spreadSheet.createRow(rowid++);
			List<String> lines = (List<String>) valores.get(indice);
			int cellId = 0;
			for (String line : lines) {
				Cell cell = row.createCell(cellId++);
				cell.setCellValue(line);
			}
		}
	}

	public static void exportaArquivo(String nomeArquivo) {
		System.out.println("Escrevendo o workbook no novo arquivo");
		valoresPlanilha();
		System.out.println("Formatando planilha");
		formataPlanilha();
		try {
			FileOutputStream planilha = new FileOutputStream(new File(nomeArquivo+".xlsx"));
			workbook.write(planilha);
			planilha.close();
			workbook.close();
			System.out.println("Sucesso na criacao do arquivo!");
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
	
	public static void formataPlanilha() {
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GOLD.getIndex());
		spreadSheet.setDefaultColumnStyle(0, style);
		spreadSheet.setDefaultColumnWidth(50);
		
	}

}
