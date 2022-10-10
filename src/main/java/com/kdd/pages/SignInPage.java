package com.kdd.pages;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.pagefactory.AjaxElementLocatorFactory;

import com.kdd.config.DriverManager;
import com.kdd.utility.ElementOperations;

public class SignInPage extends ElementOperations {
	
	private WebDriver driver;
	
	public SignInPage() {
		driver = DriverManager.getInstance().getDriver();
		PageFactory.initElements(new AjaxElementLocatorFactory(driver, implicitWaitTime), this);
	}
	
	//@FindBy(name = "username")
	@FindBy(name = "userName")
	private WebElement userNameTextbox;
	
	//@FindBy(name = "password")
	@FindBy(name = "password")
	private WebElement passwordTextbox;

	//@FindBy(css = ".button-bar>.button")
	@FindBy(name = "submit")
	private WebElement loginButton;
	
	//@FindBy(css = "#WelcomeContent")
	@FindBy(xpath = "/html/body/div[2]/table/tbody/tr/td[2]/table/tbody/tr[4]/td/table/tbody/tr/td[2]/table/tbody/tr[1]/td/h3")
	private WebElement welcomeText;
	
	@FindBy (linkText = "Register Now!")
	private WebElement registerLink;
	
	@FindBy(linkText = "SIGN-OFF")
	private WebElement logoutLink;
	
	public void enterUserName(String userName) {
		enterText(userNameTextbox, userName);
	}
	
	public void enterPassword(String password) {
		enterText(passwordTextbox, password);
	}
	
	public void clickLoginButton() {
		loginButton.click();
	}
	public void clickLogoutLink() {
		logoutLink.click();
	}
	
	public String getWelcomeText() {
		return getElementText(welcomeText);
	}
	
	public boolean isSignOutLinkDisplayed() {
		return isElementDisplayed(logoutLink);
	}
}
