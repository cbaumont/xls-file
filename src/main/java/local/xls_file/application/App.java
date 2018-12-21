package local.xls_file.application;

import local.xls_file.Utils.XlsUtils;
import local.xls_file.services.Service;

public class App {
	public static void main(String[] args) {
		Service service = new Service();
		service.entidadesXls();
		XlsUtils.exportaArquivo();

	}
}
