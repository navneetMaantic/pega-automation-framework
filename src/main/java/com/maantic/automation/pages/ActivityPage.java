package com.maantic.automation.pages;

import com.maantic.automation.base.BasePage;
import com.maantic.automation.utils.CommonUtils;
import org.openqa.selenium.By;

public class ActivityPage extends BasePage {

    private By iframe01_name = By.xpath("//iframe[@name='PegaGadget1Ifr']");
    private By btn_Actions = By.xpath("//button[text()='Actions']");
    private By btn_ActionsRun = By.xpath("(//span[text()='Run'])[2]");
    private By btn_RunAgain = By.xpath("//div[contains(text(),'Run')]");
    private By txt_RunResult = By.xpath("//div[@class='error']/h1");
    private By btnClose = By.xpath("//button[@title='Close']");

//    private By getRuleName(String ruleName){
//        return By.xpath("//span[@title='Purpose'][contains(text(),'"+ruleName+"')]");
//    }
//
//    public boolean isDecisionTableDisplayed(String ruleName) throws InterruptedException {  //check if Decision Table is displayed after clicking on search results
//        Thread.sleep(5000);
//        //CommonUtils.switchToIframe(iframe01_name);
//        return  CommonUtils.isElementPresent(getRuleName(ruleName));
//    }

    public void clickActionsRunOfActivity() throws InterruptedException {
        CommonUtils.click(btn_Actions);
        Thread.sleep(3000);
        CommonUtils.click(btn_ActionsRun);
    }

    public String switchWindowActivity() throws InterruptedException {   //switches window to run Activity & fetch results using another function
        String HandleBefore = getDriver().getWindowHandle();
        for(String winHandle:getDriver().getWindowHandles()) {
            getDriver().switchTo().window(winHandle);
        }
        Thread.sleep(5000);
        String output = runActivityWithoutParameter();
        Thread.sleep(5000);
        getDriver().close();
        getDriver().switchTo().window(HandleBefore);
        getDriver().switchTo().defaultContent();
        CommonUtils.switchToIframe(iframe01_name);
        return output;
    }

    public String runActivityWithoutParameter() throws InterruptedException {   //run activity and get result
        CommonUtils.click(btn_RunAgain);
        Thread.sleep(5000);
        String HandleBefore = getDriver().getWindowHandle();
        for(String winHandle:getDriver().getWindowHandles()) {
            getDriver().switchTo().window(winHandle);
        }
        String act_Output = CommonUtils.getElementText(txt_RunResult);
        System.out.println("Actual Result: "+act_Output);
        getDriver().close();
        getDriver().switchTo().window(HandleBefore);
        return act_Output;
    }
    
    public void closeActivity() throws InterruptedException {
    	CommonUtils.click(btnClose);
    	Thread.sleep(5000);
    }
}
