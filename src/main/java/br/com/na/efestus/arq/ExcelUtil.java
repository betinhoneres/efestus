/**
 *  
 */

package br.com.na.efestus.arq;


import java.io.File;
import java.util.ArrayList;
import java.util.Date;

import jxl.Cell;
import jxl.CellType;
import jxl.DateCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;

/**
 * Classe que realiza a leitura de dados de uma planilha na extensão Xls.
 * 
 * @since 04/09/2014
 */
public class ExcelUtil{
	
	private ArrayList<String>lstVariaveis;
	private ArrayList<String>lstNomeCT;
	private static String path;
	private static String casoTeste;
	private static boolean pathNovo = false;
	DataUtil dataUtil = new DataUtil();

	public static void setPath(String pathAtual) {
		if(!pathAtual.equals(path)  ){
			ExcelUtil.pathNovo = true;
		}
		ExcelUtil.path = pathAtual;
	}

	public static void setCasoTeste(String casoTeste) {
		ExcelUtil.casoTeste = casoTeste;
	}
	
	/**
	 * Método que se passa o nome da linha e o nome da coluna, 
	 * e o método retorna o valor.
	 * Obs.: O valor da coluna e da linha passados deverão estar na primeira linha
	 * e primeira coluna da Planilha.
	 * 
	 * Ex.: Lerplanilha("CT-013", "Descricao");
	 * 
	 * @param nomeCT  - Nome do Caso de Teste.
	 * @param nomeColuna - Nome da Coluna.
	 * @return String - Retorna o valor da célula.
	 * 
	 * @since 04/09/2014
	 * @throws Exception 
	 */
	public String lerPlanilha(String nomeColuna) throws Exception {
		int abaPlan = 1;
		int linhaBase = 9;
		int colunaBase = 0;
		int qtdLinhasAceitaveis = 1000;
		try { 
			File fp = new File(path);
			WorkbookSettings conf = new WorkbookSettings();  
			conf.setEncoding("ISO-8859-1"); 
			Workbook wb = Workbook.getWorkbook(fp, conf);   
			Sheet aba = wb.getSheet(abaPlan); 
			Cell cellValue;
			if(aba.getRows() > qtdLinhasAceitaveis){
				throw new Exception("Quantidade de linhas na planilha maior que 1000!");
			}
			if(pathNovo || lstVariaveis == null){
				lstVariaveis = new ArrayList<String>();
				for (int i = 0; i < aba.getColumns(); i++) {
					lstVariaveis.add(aba.getCell(i, linhaBase).getContents().toString().trim());
				}
				pathNovo =false;
			}
			if(pathNovo || lstNomeCT == null){
				lstNomeCT = new ArrayList<String>();
				for (int i = 0; i < aba.getRows(); i++) {
					lstNomeCT.add(aba.getCell(colunaBase, i).getContents().toString().trim());
				}
				pathNovo =false;
			}
			cellValue = aba.getCell(lstVariaveis.indexOf(nomeColuna), lstNomeCT.indexOf(casoTeste));
			if (cellValue.getType().equals(CellType.DATE)){  
				return lerDataExcel(cellValue);
			}else{
				return cellValue.getContents().toString().trim();
			}
		}   
		catch(ArrayIndexOutOfBoundsException ioe) {   
			throw new Exception("casoTeste: -> " + casoTeste + 
					" <- Ou Coluna: -> " + nomeColuna + " <- não foram encontrados na Planilha!!!  " + ioe);  
		}
	}
	
	public String[] lerPlanColuna1() throws Exception {
		File fp = new File(path);  
		int abaPlan = 1;
		int colunaBase = 0;
		int qtdLinhasAceitaveis = 1000;
		StringBuilder parametros = new StringBuilder(); 

		try { 
			WorkbookSettings conf = new WorkbookSettings();  
			conf.setEncoding("ISO-8859-1"); 
			Workbook wb = Workbook.getWorkbook(fp, conf);   
			Sheet aba = wb.getSheet(abaPlan); 
			
			if(aba.getRows() > qtdLinhasAceitaveis){
				throw new Exception("Quantidade de linhas na planilha maior que 1000!");
			}
 			for (int i = 0; i < aba.getRows(); i++) {
				String valor = (aba.getCell(colunaBase, i).getContents().toString().trim());
				if(valor.contains("Elaborado por")){
					parametros.append(valor);
					parametros.append("\n");
					for (int j = 0; j < aba.getColumns(); j++) {
						String valor2 = (aba.getCell(j, i).getContents().toString().trim());
						if(valor2.contains("Arquivo")){
							parametros.append(valor2);
						} 
					}
				}
			}
 			String[] parametros1 = parametros.toString().split("\n");
			if(parametros1.length == 2){
				return parametros1;
			}else{
				throw new ArrayIndexOutOfBoundsException();
			}
		}   
		catch(ArrayIndexOutOfBoundsException ioe) {   
			throw new Exception("Os valores dos parâmetros: 'Elaborado por' ou 'Arquivo' não foram encontrados na Planilha!  " + ioe);  
		}
	}
	
	public String[] lerPlanLinha4() throws Exception {
		File fp = new File(path);  
		int abaPlan = 1;
		int linhaBase = 3;
		int qtdLinhasAceitaveis = 1000;

		try { 
			WorkbookSettings conf = new WorkbookSettings();  
			conf.setEncoding("ISO-8859-1"); 
			Workbook wb = Workbook.getWorkbook(fp, conf);   
			Sheet aba = wb.getSheet(abaPlan); 
			StringBuilder parametros = new StringBuilder();
			
			if(aba.getRows() > qtdLinhasAceitaveis){
				throw new Exception("Quantidade de linhas na planilha maior que 1000!");
			}
			for (int i = 0; i < aba.getColumns(); i++) {
				String valor = (aba.getCell(i, linhaBase).getContents().toString().trim());
				if(valor.contains("Sistema")){
					parametros.append(valor);
					parametros.append("\n");
				}else if(valor.contains("Subsistema")){
					parametros.append(valor);
					parametros.append("\n");
				}else if(valor.contains("Módulo")){
					parametros.append(valor);
					parametros.append("\n");
				}else if(valor.contains("Caso de Uso / Requisito")){
					parametros.append(valor);
				}
			}
			String[] parametros1 = parametros.toString().split("\n");
			if(parametros1.length == 4){
				return parametros1;
			}else{
				throw new ArrayIndexOutOfBoundsException();
			}
		}   
		catch(ArrayIndexOutOfBoundsException ioe) {   
			throw new Exception("Algum valor da linha 4 não foi encontrado na Planilha! Método lerPlanLinha4  " + ioe);  
		}
	}
	
	public String lerPlanLinha5() throws Exception {
		File fp = new File(path);  
		int abaPlan = 1;
		int linhaBase = 4;
		int qtdLinhasAceitaveis = 1000;

		try { 
			WorkbookSettings conf = new WorkbookSettings();  
			conf.setEncoding("ISO-8859-1"); 
			Workbook wb = Workbook.getWorkbook(fp, conf);   
			Sheet aba = wb.getSheet(abaPlan); 

			if(aba.getRows() > qtdLinhasAceitaveis){
				throw new Exception("Quantidade de linhas na planilha maior que 1000!");
			}
			for (int i = 0; i < aba.getColumns(); i++) {
				String parametro = (aba.getCell(i, linhaBase).getContents().toString().trim());
				if(parametro.contains("Código / Projeto")){
					return parametro;
				}
			}
		}   
		catch(ArrayIndexOutOfBoundsException ioe) {   
			throw new Exception("A coluna não foi encontrada na Planilha!!!  " + ioe);  
		}
		return null;
	}
	
	public String lerPlanLinha6() throws Exception {
		File fp = new File(path);  
		int abaPlan = 1;
		int linhaBase = 5;
		int qtdLinhasAceitaveis = 1000;

		try { 
			WorkbookSettings conf = new WorkbookSettings();  
			conf.setEncoding("ISO-8859-1"); 
			Workbook wb = Workbook.getWorkbook(fp, conf);   
			Sheet aba = wb.getSheet(abaPlan); 

			if(aba.getRows() > qtdLinhasAceitaveis){
				throw new Exception("Quantidade de linhas na planilha maior que 1000!");
			}
			for (int i = 0; i < aba.getColumns(); i++) {
				String parametro = (aba.getCell(i, linhaBase).getContents().toString().trim());
				if(parametro.contains("Elemento sob Teste")){
					return parametro;
				}
			}
		}   
		catch(ArrayIndexOutOfBoundsException ioe) {   
			throw new Exception("A coluna não foi encontrada na Planilha!!!  " + ioe);  
		}
		return null;
	}
	
	public String lerProcedimentoTestePlan(String nomeCasoTeste) throws Exception {
		File fp = new File(path);  
		int abaPlan = 1;
		int colunaBaseCT = 0;
		int colunaBaseProcedimentoTeste = 2;
		int qtdLinhasAceitaveis = 1000;

		try { 
			WorkbookSettings conf = new WorkbookSettings();  
			conf.setEncoding("ISO-8859-1"); 
			Workbook wb = Workbook.getWorkbook(fp, conf);   
			Sheet aba = wb.getSheet(abaPlan); 
			Cell cellValue;

			if(aba.getRows() > qtdLinhasAceitaveis){
				throw new Exception("Quantidade de linhas na planilha maior que 1000!");
			}
			if(lstNomeCT == null){
				lstNomeCT = new ArrayList<String>();
				for (int i = 0; i < aba.getRows(); i++) {
					lstNomeCT.add(aba.getCell(colunaBaseCT, i).getContents().toString().trim());
				}
			}
			cellValue = aba.getCell(colunaBaseProcedimentoTeste, lstNomeCT.indexOf(nomeCasoTeste));
			return cellValue.getContents().toString().trim();
		}   
		catch(ArrayIndexOutOfBoundsException ioe) {   
			throw new Exception("casoTeste: -> " + nomeCasoTeste + 
					" <- Ou Coluna: -> " + colunaBaseProcedimentoTeste + " <- não foram encontrados na Planilha!!!  " + ioe);  
		}
	}
	
	/**
	 * Método que rotorna uma String para quando o valor da célula é uma Data no Excel.
	 * 
	 * Ex.: lerDataExcel(cellValue);
	 * 
	 * @param cellValue
	 * @return String
	 * @throws Exception 
	 */
	private String lerDataExcel(Cell cellValue) throws Exception{
		DateCell dateCell = (DateCell) cellValue;
		Date data = dateCell.getDate();
		return dataUtil.formataData(data);
	}
	
}

