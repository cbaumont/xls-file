package local.xls_file.services;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import local.xls_file.entidades.Entidade;
import local.xls_file.utils.XlsUtils;

public class Service {

	private List<Entidade> entidades = new ArrayList<Entidade>();

	public void adicionaEntidades() {
		System.out.println("Criando entidades");
		for (int i = 0; i < 10; i++) {
			Entidade entidade = new Entidade();
			entidade.setAtributo1("Valor 01 " + entidade.hashCode());
			entidade.setAtributo2("Valor 02 " + entidade.hashCode());
			entidade.setAtributo3("Valor 03 " + entidade.hashCode());
			entidades.add(entidade);
		}

	}

	public void entidadesXls() {
		System.out.println("Gravando valores no workbook");
		adicionaEntidades();
		Map<String, List<String>> cabecalho = new LinkedHashMap<String, List<String>>();
		cabecalho.put(cabecalho.hashCode()+"", Arrays.asList("Coluna 01", "Coluna 02", "Coluna 03"));
		Map<String, List<String>> linha = new LinkedHashMap<String, List<String>>();
		for (Entidade entidade : entidades) {
			linha.put(entidade.hashCode() + "", Arrays.asList(entidade.getAtributo1(), 
					entidade.getAtributo2(), entidade.getAtributo3()));
		}
		XlsUtils.geraCabecalho(cabecalho);
		XlsUtils.geraLinhas(linha);
	}

}
