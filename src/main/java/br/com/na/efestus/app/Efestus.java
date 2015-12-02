package br.com.na.efestus.app;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

import br.com.na.efestus.arq.LerXls;
import br.com.na.efestus.arq.ParametrossUtil;
import br.com.na.efestus.pageobjects.LoginPage;
import br.com.na.efestus.pageobjects.NovoPost;

public class Efestus {

	public static void main(String[] args) {

		WebDriver driver = new FirefoxDriver();
		
		LoginPage login = new LoginPage(driver);
		
		try {
			
			ParametrossUtil.setArquivoProperties("parametros.properties");
			
			login.digitarLogin(ParametrossUtil.getValueAsString("usuario"));
			login.digitarSenha(ParametrossUtil.getValueAsString("senha"));
			login.acessarPagina();
			
			//HomePage homePage = new HomePage(driver);
			//homePage.todosOsPosts();

			LerXls xls = new LerXls();
			xls.setPath(ParametrossUtil.getValueAsStringOriginal("planilha"));
			
			for(int linha = 1; !xls.lerLinhaColunaPlan(linha, 1).equalsIgnoreCase("x"); linha++){
				
				NovoPost post = new NovoPost(driver);
				post.adicionarTitulo(xls.lerLinhaColunaPlan(linha, 1));
				post.preencherCorpoDoPost(xls.lerLinhaColunaPlan(linha, 2));
				post.salvarRescunho();
				
			}			
			
		} catch (Exception e) {
			e.printStackTrace();
			driver.close();
			driver = null;
		}
		
	}
	
}