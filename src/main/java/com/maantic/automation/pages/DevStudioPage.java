package com.maantic.automation.pages;

import com.maantic.automation.base.BasePage;
import com.maantic.automation.utils.CommonUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;

import java.util.List;

public class DevStudioPage extends BasePage {

    private By iframe01_name = By.xpath("//iframe[@name='PegaGadget0Ifr']");
    private By iframe02_name = By.xpath("//iframe[@name='PegaGadget1Ifr']");
//    private By search_txtBox = By.xpath("(//input[contains(@name, 'SearchText')])[1]");
//    private By search_icon_btn = By.xpath("//button[@title='Search'][i]");
    private By lblRecords = By.xpath("//div[@title='Records']");
    private By lblDecision = By.xpath("//span[text()='Decision']");
    private By lblDecisionTable = By.xpath("(//a[contains(text(),'Decision Table')])[1]");
    private By filterPurposeDecisionTable = By.xpath("//div[contains(text(),'Purpose')]/following::a[1]");
    private By txtFilterDecisionTable = By.xpath("//label[text()='Search Text']/following::input");
    private By btnFilterApply = By.xpath("//label[text()='Search Text']/following::button[1]");
    private By filterRulesetVersion = By.xpath("//div[contains(text(),'RuleSet:Version')]/following::a[1]");
    private By txtFilterRulesetVersion = By.xpath("//label[text()='Search Text']/following::input");
    private By imgNoItems = By.xpath("//i[@title='No data ']");
    private By lblProcess = By.xpath("//span[text()='Process']");
    private By lblSLA = By.xpath("(//a[contains(text(),'Service Level Agreement')])[1]");
    private By filterSLA = By.xpath("//div[contains(text(),'Service Level')]/following::a[1]");
    private By txtFilterSLA = By.xpath("//label[text()='Search Text']/following::input");
    private By lblTechnical = By.xpath("//span[text()='Technical']");
    private By lblActivity = By.xpath("(//a[contains(text(),'Activity')])[1]");
    private By filterActivityName = By.xpath("//div[contains(text(),'Activity Name')]/following::a[1]");
    private By txtFilterActivity = By.xpath("//label[text()='Search Text']/following::input");


    private By getRowOfRuletype(String ruleName){
        return By.xpath("//div[contains(text(),'"+ruleName+"')]");
    }
//    private By getTdDecisionTableName(String className, String ruleName){
//        return By.xpath("//table[@id='bodyTbl_right']/tbody/tr/td/div/span[text()='"+ className +"']/preceding::tr[1]/td/nobr/span/a[text()='"+ ruleName +"']");
//    }
//    private By getTdSLAName(String className, String ruleName){
//        return By.xpath("//table[@id='bodyTbl_right']/tbody/tr/td/div/span[text()='"+ className +"']/preceding::tr[1]/td/nobr/span/a[text()='"+ ruleName +"']");
//    }
//    private By getTdRuleSetVersion(String ruleSetVersion){
//        return By.xpath("(//table[@id='bodyTbl_right'])[2]/tbody/tr/td/div[contains(text(),'" + ruleSetVersion + "')]");
//    }

//    public void selectCorrectRule(String ruleName, String className){   //to select out correct rule row from the search result
//        List<WebElement> rowList = getDriver().findElements(By.xpath("(//table[@id='bodyTbl_right'])[1]/tbody/tr/td[3]/div"));
//        List<WebElement> rowList2 = getDriver().findElements(By.xpath("(//table[@id='bodyTbl_right'])[1]/tbody/tr/td[4]/div"));
//        int i;
//        for (i=0; i<rowList.size();i++){
//            String lblruleName = rowList.get(i).getText();
//            String lblclassName = rowList2.get(i).getText();
//            if(lblruleName.equalsIgnoreCase(ruleName) && lblclassName.equalsIgnoreCase(className)){
//                getDriver().findElement(By.xpath("//a[text()='"+ruleName+"']/preceding::td/span[@class='expandRowDetails'][1]")).click();
//                break;
//            }
//        }
//    }

//    public void enterSearchTermInSearchBox(String item) {    //enter search text in searchBox
//        CommonUtils.waitForVisibilityOfElement(search_txtBox);
//        CommonUtils.enterText(search_txtBox, item);
//    }

//    public void clickOnSearchIcon() {    //click on Search button
//        CommonUtils.waitForVisibilityOfElement(search_icon_btn);
//        CommonUtils.click(search_icon_btn);
//    }
//    public void selectDropdownContains(){   //to select 'exact match' from dropdown
//        try {
//            Thread.sleep(3000);
//        } catch (InterruptedException e) {
//            throw new RuntimeException(e);
//        }
//        Select se = new Select(getDriver().findElement(By.xpath("//select[contains(@name,'Method')]")));
//        se.selectByIndex(2);
//    }
    //click on the search result after verifying ruleName, ruleType & ruleSetVersion
//    public Boolean clickSearchResults(String ruleType, String className, String ruleSetVersion, String ruleName) throws InterruptedException {
//        Boolean temp = false;
//        if (ruleType.equalsIgnoreCase("Decision_Table")) {
//            if (CommonUtils.isElementPresent(getTdDecisionTableName(className, ruleName))) {
//                if (CommonUtils.getElementText(getTdDecisionTableName(className, ruleName)).equalsIgnoreCase(ruleName)) {
//                    CommonUtils.waitForVisibilityOfElement(search_txtBox);
//                    selectCorrectRule(ruleName, className);
//                    Thread.sleep(3000);
//                    CommonUtils.click(getTdRuleSetVersion(ruleSetVersion));
//                    Thread.sleep(5000);
//                    temp = true;
//                }
//            }
//        } else if (ruleType.equalsIgnoreCase("Activity")) {
//
//        } else if (ruleType.equalsIgnoreCase("SLA")) {
//            if (CommonUtils.isElementPresent(getTdSLAName(className, ruleName))) {
//                if (CommonUtils.getElementText(getTdSLAName(className, ruleName)).equalsIgnoreCase(ruleName)) {
//                    CommonUtils.waitForVisibilityOfElement(search_txtBox);
//                    selectCorrectRule(ruleName, className);
//                    Thread.sleep(3000);
//                    CommonUtils.click(getTdRuleSetVersion(ruleSetVersion));
//                    Thread.sleep(5000);
//                    temp = true;
//                }
//            }
//        }
//        return temp;
//    }

    public Boolean selectRuleType(int count, String ruleType, String ruleName, String rulesetVersion) throws InterruptedException {
        Boolean temp = false;
        if(count == 0) {
        	CommonUtils.click(lblRecords);
        }
        Thread.sleep(5000);
        if(ruleType.equalsIgnoreCase("Decision_Table")){
        	if(count == 0) {
        		CommonUtils.click(lblDecision);
                Thread.sleep(5000);
                CommonUtils.click(lblDecisionTable);
                Thread.sleep(5000);	
        	}
        	getDriver().switchTo().defaultContent();
            CommonUtils.switchToIframe(iframe01_name);
            CommonUtils.click(filterPurposeDecisionTable);
            Thread.sleep(3000);
            CommonUtils.enterText(txtFilterDecisionTable, ruleName);
            Thread.sleep(2000);
            CommonUtils.click(btnFilterApply);
            Thread.sleep(3000);
            if(!CommonUtils.isElementPresent(imgNoItems)){
            	CommonUtils.click(filterRulesetVersion);
                Thread.sleep(2000);
                CommonUtils.enterText(txtFilterRulesetVersion, rulesetVersion);
                CommonUtils.click(btnFilterApply);
                Thread.sleep(2000);
                if(!CommonUtils.isElementPresent(imgNoItems)){
                    CommonUtils.click(getRowOfRuletype(ruleName));
                    Thread.sleep(8000);
                    getDriver().switchTo().defaultContent();
                    CommonUtils.switchToIframe(iframe02_name);
                    temp = true;
                }
            }
            
        } else if (ruleType.equalsIgnoreCase("SLA")) {
            if(count == 0) {
            	CommonUtils.click(lblProcess);
                Thread.sleep(5000);
                CommonUtils.click(lblSLA);
                Thread.sleep(5000);
            }
            getDriver().switchTo().defaultContent();
            CommonUtils.switchToIframe(iframe01_name);
        	CommonUtils.click(filterSLA);
            Thread.sleep(3000);
            CommonUtils.enterText(txtFilterSLA, ruleName);
            Thread.sleep(2000);
            CommonUtils.click(btnFilterApply);
            Thread.sleep(3000);
            if(!CommonUtils.isElementPresent(imgNoItems)){
            	CommonUtils.click(filterRulesetVersion);
                Thread.sleep(2000);
                CommonUtils.enterText(txtFilterRulesetVersion, rulesetVersion);
                CommonUtils.click(btnFilterApply);
                Thread.sleep(2000);
                if(!CommonUtils.isElementPresent(imgNoItems)){
                    CommonUtils.click(getRowOfRuletype(ruleName));
                    Thread.sleep(12000);
                    getDriver().switchTo().defaultContent();
                    CommonUtils.switchToIframe(iframe02_name);
                    temp = true;
                }
            }
            
        } else if (ruleType.equalsIgnoreCase("Activity")) {
        	if(count ==0) {
        		CommonUtils.click(lblTechnical);
                Thread.sleep(5000);
                CommonUtils.click(lblActivity);
                Thread.sleep(5000);                	
        	}
        	getDriver().switchTo().defaultContent();
            CommonUtils.switchToIframe(iframe01_name);
            CommonUtils.click(filterActivityName);
            Thread.sleep(3000);
            CommonUtils.enterText(txtFilterActivity, ruleName);
            Thread.sleep(2000);
            CommonUtils.click(btnFilterApply);
            Thread.sleep(3000);
            if(!CommonUtils.isElementPresent(imgNoItems)){
            	CommonUtils.click(filterRulesetVersion);
                Thread.sleep(2000);
                CommonUtils.enterText(txtFilterActivity, rulesetVersion);
                CommonUtils.click(btnFilterApply);
                Thread.sleep(2000);
                if(!CommonUtils.isElementPresent(imgNoItems)){
                    CommonUtils.click(getRowOfRuletype(ruleName));
                    Thread.sleep(8000);
                    getDriver().switchTo().defaultContent();
                    CommonUtils.switchToIframe(iframe02_name);
                    temp = true;
                }
            }
            
        }
        return temp;
    }
}
