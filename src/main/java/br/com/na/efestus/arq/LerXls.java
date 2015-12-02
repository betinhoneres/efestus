package br.com.na.efestus.arq;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import jxl.Cell;
import jxl.CellType;
import jxl.DateCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;

/**
 * Classe que realizar a leitura de dados de uma planilha na extensão Xls.
 * 
 */
public class LerXls {

	private ArrayList<String> lstVariaveis;
	private ArrayList<String> lstNomeCT;
	private String path;
	private String casoTeste;
	boolean pathNovo = false;
	DataUtil dataUtil = new DataUtil();
	ExcelUtil excelUtil = new ExcelUtil();

	public void setPath(String pathAtual) {
		if (!pathAtual.equals(path)) {
			pathNovo = true;
		}
		this.path = pathAtual;
		ExcelUtil.setPath(pathAtual);
	}

	public void setCasoTeste(String casoTeste) {
		this.casoTeste = casoTeste;
		ExcelUtil.setCasoTeste(casoTeste);
	}

	/**
	 * Método que se passa o nome da linha e o nome da coluna, e o método
	 * retorna o valor. Obs.: O valor da coluna e da linha passados deverão
	 * estar na primeira linha e primeira coluna da Planilha.
	 * 
	 * Ex.: Lerplanilha("CT-013", "Descricao");
	 * 
	 * @param nomeCT
	 *            - Nome do Caso de Teste.
	 * @param nomeColuna
	 *            - Nome da Coluna.
	 * @return String - Retorna o valor da célula.
	 * 
	 * @throws Exception
	 */
	public String lerPlanilha(String nomeColuna) throws Exception {
		int abaPlan = 1;
		int linhaBase = 6;
		int colunaBase = 0;
		int qtdLinhasAceitaveis = 1000;
		try {
			File fp = new File(path);
			WorkbookSettings conf = new WorkbookSettings();
			conf.setEncoding("ISO-8859-1");
			Workbook wb = Workbook.getWorkbook(fp, conf);
			Sheet aba = wb.getSheet(abaPlan);
			Cell cellValue;
			if (aba.getRows() > qtdLinhasAceitaveis) {
				throw new Exception(
						"Quantidade de linhas na planilha maior que 1000!");
			}
			if (pathNovo || lstVariaveis == null) {
				lstVariaveis = new ArrayList<String>();
				for (int i = 0; i < aba.getColumns(); i++) {
					lstVariaveis.add(aba.getCell(i, linhaBase).getContents()
							.toString().trim());
				}
				pathNovo = false;
			}
			if (pathNovo || lstNomeCT == null) {
				lstNomeCT = new ArrayList<String>();
				for (int i = 0; i < aba.getRows(); i++) {
					lstNomeCT.add(aba.getCell(colunaBase, i).getContents()
							.toString().trim());
				}
				pathNovo = false;
			}
			cellValue = aba.getCell(lstVariaveis.indexOf(nomeColuna),
					lstNomeCT.indexOf(casoTeste));
			if (cellValue.getType().equals(CellType.DATE)) {
				return lerDataExcel(cellValue);
			} else {
				return cellValue.getContents().toString().trim();
			}
		} catch (ArrayIndexOutOfBoundsException ioe) {
			throw new Exception("casoTeste: -> " + casoTeste
					+ " <- Ou Coluna: -> " + nomeColuna
					+ " <- não foram encontrados na Planilha!!!  " + ioe);
		}
	}

	public String lerPlanCasoUso() throws Exception {
		File fp = new File(path);
		int abaPlan = 1;
		int linhaBase = 3;
		int qtdLinhasAceitaveis = 1000;

		try {
			WorkbookSettings conf = new WorkbookSettings();
			conf.setEncoding("ISO-8859-1");
			Workbook wb = Workbook.getWorkbook(fp, conf);
			Sheet aba = wb.getSheet(abaPlan);

			if (aba.getRows() > qtdLinhasAceitaveis) {
				throw new Exception(
						"Quantidade de linhas na planilha maior que 1000!");
			}
			for (int i = 0; i < aba.getColumns(); i++) {
				String casoUso = (aba.getCell(i, linhaBase).getContents()
						.toString().trim());
				if (casoUso.contains("Caso de Uso / Requisito:")) {
					return casoUso;
				}
			}
		} catch (ArrayIndexOutOfBoundsException ioe) {
			throw new Exception("A coluna não foi encontrada na Planilha!!!  "
					+ ioe);
		}
		return null;
	}

	/**
	 * Método que retorna uma lista com os valores de acordo com o número da
	 * linha.
	 * 
	 * @param lin
	 *            - Linha
	 * @return ArrayList<String> - Uma lista com os valores da linha passada.
	 * @throws Exception
	 */
	public ArrayList<String> lerLinhaPlanilha(int lin) throws Exception {
		ArrayList<String> lista = new ArrayList<String>();
		File fp = new File(path);
		try {
			Workbook wb = Workbook.getWorkbook(fp);
			Sheet aba = wb.getSheet(0);
			Cell cellValue = null;

			for (int i = 0; i < aba.getColumns(); i++) {
				cellValue = aba.getCell(i, lin - 1);
				// Se a célula conter data, trata o valor e adiciona na lista.
				if (cellValue.getType().equals(CellType.DATE)) {
					lista.add(lerDataExcel(cellValue));
				} else {
					lista.add(cellValue.getContents().toString());
				}
			}
			return lista;
		} catch (Exception ioe) {
			ioe.printStackTrace();
		}
		return lista;
	}

	/**********************************************************************
	 * inserirStringPlan(Param1, Param2, Param3): inclui dados em uma planilha.
	 * Onde:
	 * 
	 * Param1 = Linha Param2 = Coluna Param3 = Conteúdo
	 * 
	 */
	@SuppressWarnings("deprecation")
	public void inserirStringPlan(int lin, int col, String conteudo) {
		try {
			int abaPlan = 1;
			col = col - 1;
			lin = lin - 1;
			HSSFWorkbook plan = new HSSFWorkbook(new FileInputStream(path));
			HSSFSheet aba = plan.getSheetAt(abaPlan);
			HSSFRow linhaFolha = aba.getRow(lin);
			HSSFCell celula = linhaFolha.getCell((short) col);
			celula.setCellType(HSSFCell.CELL_TYPE_STRING);
			celula.setCellValue(conteudo);
			FileOutputStream stream = new FileOutputStream(path);
			plan.write(stream);
			stream.flush();
			stream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**********************************************************************
	 * inserirStringPlan(Param1, Param2, Param3): inclui dados em uma planilha.
	 * Onde:
	 * 
	 * Param1 = Linha Param2 = Coluna Param3 = Conteúdo
	 * 
	 */
	@SuppressWarnings("deprecation")
	public void inserirIntPlan(int lin, int col, int conteudo) {
		try {
			int abaPlan = 1;
			col = col - 0;
			lin = lin - 1;
			HSSFWorkbook plan = new HSSFWorkbook(new FileInputStream(path));
			HSSFSheet aba = plan.getSheetAt(abaPlan);
			HSSFRow linhaFolha = aba.getRow(lin);
			HSSFCell celula = linhaFolha.getCell((short) col);
			celula.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
			celula.setCellValue(conteudo);
			FileOutputStream stream = new FileOutputStream(path);
			plan.write(stream);
			stream.flush();
			stream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * @param lin
	 * @param col
	 * @return string
	 */
	public String lerLinhaColunaPlan(int lin, int col) {
		File fp = new File(path);
		try {
			Workbook wb = Workbook.getWorkbook(fp);
			Sheet aba = wb.getSheet(0);
			Cell cellValue = aba.getCell(col - 1, lin - 1);
			if (aba.getCell(col - 1, lin - 1).getType().equals(CellType.DATE)) {
				return lerDataExcel(cellValue);
			}
			return cellValue.getContents().toString();
		} catch (Exception ioe) {
			ioe.printStackTrace();
		}
		return null;
	}

	/**
	 * @param lin
	 * @param col
	 * @return string
	 */
	public String lerLinhaColunaPlan(int lin, int col, int numAba) {
		File fp = new File(path);
		try {
			Workbook wb = Workbook.getWorkbook(fp);
			Sheet aba = wb.getSheet(numAba);
			Cell cellValue = aba.getCell(col - 1, lin - 1);
			if (aba.getCell(col - 1, lin - 1).getType().equals(CellType.DATE)) {
				return lerDataExcel(cellValue);
			}
			return cellValue.getContents().toString();
		} catch (Exception ioe) {
			ioe.printStackTrace();
		}
		return null;
	}

	/**
	 * Método que rotorna uma String para quando o valor da célula é uma Data no
	 * Excel.
	 * 
	 * Ex.: lerDataExcel(cellValue);
	 * 
	 * @param cellValue
	 * @return String
	 * @throws Exception
	 */
	private String lerDataExcel(Cell cellValue) throws Exception {
		DateCell dateCell = (DateCell) cellValue;
		Date data = dateCell.getDate();
		return dataUtil.formataData(data);
	}

	/**********************************************************************
	 * naox(parametro): Serve para verificar se o conteúdo da
	 * variável(parâmetro) é igual a 'X'. Este artifício é usando para indicar
	 * que o campo não terá nenhuma ação.
	 * 
	 * Exemplo: if (naox(vgLink1)) {
	 * 
	 * }
	 */
	public boolean naox(String valString) {
		if (valString.equalsIgnoreCase("x")) {
			return false;
		} else
			return true;
	}

	public Integer totalDeColunas(int numeroAba){
		
		File fp = new File(path);
		WorkbookSettings conf = new WorkbookSettings();  
		conf.setEncoding("ISO-8859-1"); 
		Workbook wb;
		try {
			wb = Workbook.getWorkbook(fp, conf);
			Sheet aba = wb.getSheet(numeroAba);
			return aba.getColumns();
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}   
		return 0; 
	}

	public Integer totalDeLinhas(int numeroAba){
		
		File fp = new File(path);
		WorkbookSettings conf = new WorkbookSettings();  
		conf.setEncoding("ISO-8859-1"); 
		Workbook wb;
		try {
			wb = Workbook.getWorkbook(fp, conf);
			Sheet aba = wb.getSheet(numeroAba);
			return aba.getRows();
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}   
		return 0; 
	}
}