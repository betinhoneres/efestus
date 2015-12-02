package br.com.na.efestus.pageobjects;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

import br.com.na.efestus.arq.ParametrossUtil;
import br.com.na.efestus.arq.UtilPaginaWeb;

public class HomePage extends UtilPaginaWeb {

	private WebDriver driver;
	private String pagina = ParametrossUtil.getValueAsString("homePage");
	
	By todosOsPostsLocator = By.linkText("Posts");
    By adicionarNovoLocator = By.linkText("Adicionar Novo");
	
	public HomePage(WebDriver driver) {
		this.driver = driver;
		driver.get(pagina);
		waitForPageToLoad(driver);
	}
	
	public HomePage todosOsPosts(){
		
		driver.findElement(todosOsPostsLocator).click();;
		waitForPageToLoad(driver);
		return this;
		
	}
	
	public HomePage adicionarNovo(){

		todosOsPosts();
		driver.findElement(adicionarNovoLocator).click();
		waitForPageToLoad(driver);
		return this;
		
	}
	
}