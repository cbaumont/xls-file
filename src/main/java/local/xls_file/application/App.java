package local.xls_file.application;

import local.xls_file.services.Service;
import local.xls_file.utils.XlsUtils;

public class App {
	public static void main(String[] args) {
		Service service = new Service();
		service.entidadesXls();
		XlsUtils.exportaArquivo("Planilha teste2");

	}
}
