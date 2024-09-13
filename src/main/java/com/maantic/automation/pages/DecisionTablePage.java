package com.maantic.automation.pages;

import com.maantic.automation.base.BasePage;
import com.maantic.automation.utils.CommonUtils;
import org.openqa.selenium.By;

public class DecisionTablePage extends BasePage {

    private By iframe01_name = By.xpath("//iframe[@name='PegaGadget1Ifr']");
    private By frameDecisionTableRun = By.id("RunRuleBottom");
    private By btn_Actions = By.xpath("//button[text()='Actions']");
    private By btn_ActionsRun = By.xpath("(//span[text()='Run'])[2]");
    private By txt_newWindow_RelCode = By.id("RelCode");
    private By btn_RunAgain = By.id("RunAgainButton");
    private By txt_RunResult = By.xpath("//p[text()='Return ']/b");
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

    public void clickActionsRunOfDecisionTable() throws InterruptedException {
        CommonUtils.click(btn_Actions);
        Thread.sleep(3000);
        CommonUtils.click(btn_ActionsRun);
    }

    public String switchWindowDecisionTable(String dt_condition){   //switches window to run decision table & fetch results using another function
        String HandleBefore = getDriver().getWindowHandle();
        for(String winHandle:getDriver().getWindowHandles()) {
            getDriver().switchTo().window(winHandle);
        }
        String output = enterConditionDecisionTable(dt_condition);
        getDriver().close();
        getDriver().switchTo().window(HandleBefore);
        getDriver().switchTo().defaultContent();
        CommonUtils.switchToIframe(iframe01_name);
        return output;
    }

    public String enterConditionDecisionTable(String dt_condition){   //run decision table and get result
        CommonUtils.switchToIframe(frameDecisionTableRun);
        CommonUtils.enterText(txt_newWindow_RelCode, dt_condition);
        CommonUtils.click(btn_RunAgain);
        String dt_Output = CommonUtils.getElementText(txt_RunResult);
        System.out.println("Actual Result: "+dt_Output);
        return dt_Output;
    }
    
    public void closeDtable() throws InterruptedException {
    	CommonUtils.click(btnClose);
    	Thread.sleep(5000);
    }
}
