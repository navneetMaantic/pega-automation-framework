package com.maantic.automation.pages;

import com.maantic.automation.base.BasePage;
import com.maantic.automation.utils.CommonUtils;
import org.openqa.selenium.By;

public class SLAPage extends BasePage {

    private By iframe01_name = By.xpath("//iframe[@name='PegaGadget0Ifr']");
    private By txt_GoalDays = By.xpath("//h2[contains(text(), 'Goal')]/following::input[2]");
    private By txt_GoalHrs = By.xpath("//h2[contains(text(), 'Goal')]/following::input[3]");
    private By txt_GoalMins = By.xpath("//h2[contains(text(), 'Goal')]/following::input[4]");
    private By txt_GoalSecs = By.xpath("//h2[contains(text(), 'Goal')]/following::input[5]");
    private By txt_DeadlineDays = By.xpath("//h2[contains(text(), 'Deadline')]/following::input[2]");
    private By txt_DeadlineHrs = By.xpath("//h2[contains(text(), 'Deadline')]/following::input[3]");
    private By txt_DeadlineMins = By.xpath("//h2[contains(text(), 'Deadline')]/following::input[4]");
    private By txt_DeadlineSecs = By.xpath("//h2[contains(text(), 'Deadline')]/following::input[5]");
    private By txt_PassedDeadlineDays = By.xpath("//h2[contains(text(), 'Passed deadline')]/following::input[2]");
    private By txt_PassedDeadlineHrs = By.xpath("//h2[contains(text(), 'Passed deadline')]/following::input[3]");
    private By txt_PassedDeadlineMins = By.xpath("//h2[contains(text(), 'Passed deadline')]/following::input[4]");
    private By txt_PassedDeadlineSecs = By.xpath("//h2[contains(text(), 'Passed deadline')]/following::input[5]");
    private By btnClose = By.xpath("//button[@title='Close']");

//    private By getRuleName(String ruleName){
//        return By.xpath("//span[@title='Service Level'][contains(text(),'"+ruleName+"')]");
//    }
//
//    public boolean isSLADisplayed(String ruleName) throws InterruptedException {  //check if Decision Table is displayed after clicking on search results
//        Thread.sleep(5000);
//        CommonUtils.switchToIframe(iframe01_name);
//        return  CommonUtils.isElementPresent(getRuleName(ruleName));
//    }

    public Boolean validateSLAValues(int i, String inputParam) {
        Boolean temp = false;
        if(i == 1){
            if(CommonUtils.getAttributeValueOfElement("value", txt_GoalDays).equalsIgnoreCase(inputParam)){
                temp = true;
            }
        } else if (i == 2) {
            if(CommonUtils.getAttributeValueOfElement("value", txt_GoalHrs).equalsIgnoreCase(inputParam)){
                temp = true;
            }
        }
        else if (i == 3) {
            if(CommonUtils.getAttributeValueOfElement("value", txt_GoalMins).equalsIgnoreCase(inputParam)){
                temp = true;
            }
        }
        else if (i == 4) {
            if(CommonUtils.getAttributeValueOfElement("value", txt_GoalSecs).equalsIgnoreCase(inputParam)){
                temp = true;
            }
        }
        else if (i == 5) {
            if(CommonUtils.getAttributeValueOfElement("value", txt_DeadlineDays).equalsIgnoreCase(inputParam)){
                temp = true;
            }
        }
        else if (i == 6) {
            if(CommonUtils.getAttributeValueOfElement("value", txt_DeadlineHrs).equalsIgnoreCase(inputParam)){
                temp = true;
            }
        }
        else if (i == 7) {
            if(CommonUtils.getAttributeValueOfElement("value", txt_DeadlineMins).equalsIgnoreCase(inputParam)){
                temp = true;
            }
        }
        else if (i == 8) {
            if(CommonUtils.getAttributeValueOfElement("value", txt_DeadlineSecs).equalsIgnoreCase(inputParam)){
                temp = true;
            }
        }
        else if (i == 9) {
            if(CommonUtils.getAttributeValueOfElement("value", txt_PassedDeadlineDays).equalsIgnoreCase(inputParam)){
                temp = true;
            }
        }
        else if (i == 10) {
            if(CommonUtils.getAttributeValueOfElement("value", txt_PassedDeadlineHrs).equalsIgnoreCase(inputParam)){
                temp = true;
            }
        }
        else if (i == 11) {
            if(CommonUtils.getAttributeValueOfElement("value", txt_PassedDeadlineMins).equalsIgnoreCase(inputParam)){
                temp = true;
            }
        }
        else if (i == 12) {
            if(CommonUtils.getAttributeValueOfElement("value", txt_PassedDeadlineSecs).equalsIgnoreCase(inputParam)){
                temp = true;
            }
        }
        return temp;
    }
    
    public void closeSLAPage() throws InterruptedException {
    	CommonUtils.click(btnClose);
        Thread.sleep(5000);
    }

}
