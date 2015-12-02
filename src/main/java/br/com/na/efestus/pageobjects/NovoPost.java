package br.com.na.efestus.pageobjects;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

import br.com.na.efestus.arq.ParametrossUtil;
import br.com.na.efestus.arq.UtilPaginaWeb;

public class NovoPost extends UtilPaginaWeb {

	private WebDriver driver;
	private String pagina = ParametrossUtil.getValueAsString("novoPost");
	
	By tituloPostLocator = By.name("post_title");
    By corpoPostLocator = By.id("tinymce");
    By btnPublicarLocator = By.id("publish");
    By btnEditarAgendamentoLocator = By.className("edit-timestamp hide-if-no-js");
    By btnSalvarRescunhoLocator = By.id("save-post");
    String frameCorpoPost = "content_ifr";
    Boolean leituraDeDados = false;
    
	public NovoPost(WebDriver driver) {
		this.driver = driver;
		driver.get(pagina);
		waitForPageToLoad(driver);
	}
	
	public NovoPost adicionarTitulo(String titulo){
		driver.findElement(tituloPostLocator).sendKeys(titulo);
		return this;
	}
	
	public NovoPost preencherCorpoDoPost(String corpo){
		driver.switchTo().frame(frameCorpoPost);
		driver.findElement(corpoPostLocator).sendKeys(corpo);
		return this;
	}
	
	public NovoPost editarAgendamento(){
		driver.findElement(btnEditarAgendamentoLocator).submit();
		aguardar(2000);
		return this;
	}
	
	public NovoPost publicarOuAgendar(){
		driver.findElement(btnPublicarLocator).submit();
		waitForPageToLoad(driver);
		return this;
	}
	
	public NovoPost salvarRescunho(){
		driver.findElement(btnSalvarRescunhoLocator).submit();
		waitForPageToLoad(driver);
		return this;
	}
	
	
}