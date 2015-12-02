package br.com.na.efestus.arq;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;

public class UtilPaginaWeb {

	public void waitForPageToLoad(WebDriver driver) {

		while (!((JavascriptExecutor) driver).executeScript("return document.readyState").equals("complete")) {
			try {
				Thread.sleep(500);
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
		}
		try {
			Thread.sleep(2000);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}
	
	public void aguardar(long tempoMs){
		try {
			Thread.sleep(tempoMs);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}

}