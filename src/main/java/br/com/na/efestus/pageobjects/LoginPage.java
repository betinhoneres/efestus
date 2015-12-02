package br.com.na.efestus.pageobjects;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

import br.com.na.efestus.arq.ParametrossUtil;
import br.com.na.efestus.arq.UtilPaginaWeb;

public class LoginPage extends UtilPaginaWeb {

	private WebDriver driver;
	private String paginaInicial = ParametrossUtil.getValueAsString("paginaDeLogin");
	
	By usernameLocator = By.id("user_login");
    By passwordLocator = By.id("user_pass");
    By loginButtonLocator = By.id("wp-submit");
	
	public LoginPage(WebDriver driver) {
		this.driver = driver;
		driver.get(paginaInicial);
		waitForPageToLoad(driver);
	}
	
	public LoginPage digitarLogin(String login){
		
		driver.findElement(usernameLocator).sendKeys(login);
		return this;
	
	}
	
	public LoginPage digitarSenha(String senha){
		driver.findElement(passwordLocator).sendKeys(senha);
		return this;
	}
	
	public LoginPage acessarPagina(){
		driver.findElement(loginButtonLocator).submit();
		waitForPageToLoad(driver);
		return this;
	}
	
}