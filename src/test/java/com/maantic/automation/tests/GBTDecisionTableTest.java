package com.maantic.automation.tests;

import com.maantic.automation.base.BasePage;
import com.maantic.automation.listeners.AllureListeners;
import com.maantic.automation.pages.*;
import com.maantic.automation.utils.ExcelDataProvider;
import com.maantic.automation.utils.ExcelUtils;
import io.qameta.allure.Description;
import io.qameta.allure.Severity;
import io.qameta.allure.SeverityLevel;
import org.testng.Assert;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import java.util.Map;

@Listeners({AllureListeners.class})
public class GBTDecisionTableTest extends BasePage {

    LoginPage loginPage = new LoginPage();
    DevStudioPage homePage = new DevStudioPage();
    DecisionTablePage dtPage = new DecisionTablePage();
    String testUsername, testPassword;

    @Test(priority = 1, groups = {"Sanity", "Regression"}, description = "fetch URL", dataProviderClass = ExcelDataProvider.class, dataProvider = "ExcelTestDataLogin")
    public void getURL(Map<String, String> map) throws InterruptedException{   
    	System.out.println("Test starts");
    	appUrl = map.get("URL");
    	testUsername = map.get("Username");
    	testPassword = map.get("Password");
    }
    
    @Test(priority = 2, groups = {"Sanity", "Regression"}, description = "Verifying the Decision Table is working properly.", dataProviderClass = ExcelDataProvider.class, dataProvider = "ExcelTestDataGBT")
    @Severity(SeverityLevel.NORMAL)
    @Description("Test Case Description: Verify the Decision Table is working properly.")
    public void Decision_Table(Map<String, String> map) throws InterruptedException {
    	zoomOutChrome();
        loginPage.enterUserNameText(testUsername);
        loginPage.enterPasswordText(testPassword);
        loginPage.clickOnLogInButton();
        Thread.sleep(5000);
        int count=0;
//        homePage.selectRuleType(map.get("RuleType"), map.get("RuleName"), map.get("RuleSetVersion"));
//        homePage.enterSearchTermInSearchBox(map.get("RuleName"));
//        homePage.clickOnSearchIcon();
//        homePage.selectDropdownContains();

        //if correct ruleName present in file
//        if (homePage.clickSearchResults(map.get("RuleType"), map.get("ClassName"), map.get("RuleSetVersion"), map.get("RuleName"))) {
          if(homePage.selectRuleType(count, map.get("RuleType"), map.get("RuleName"), map.get("RuleSetVersion"))){
            dtPage.clickActionsRunOfDecisionTable();
            //now switches to decision table window
            String dt_output = dtPage.switchWindowDecisionTable(map.get("Param1"));
            //writing 'actual' result in file
            ExcelUtils.writeExcelData(dt_output, map.get("RuleType"), 17);  //'Actual' column=17
            Thread.sleep(5000);
            //writing 'pass/fail' in file
            if (map.get("Expected").equalsIgnoreCase(dt_output)) {
                ExcelUtils.writeExcelData("Pass", map.get("RuleType"), 18); //'Pass/Fail' column=18
                System.out.println("Pass");
            } else {
                ExcelUtils.writeExcelData("Fail", map.get("RuleType"), 18); //'Pass/Fail' column=18
                System.out.println("Fail");
            }
        }
        else {//if incorrect ruleName present in file
            ExcelUtils.writeExcelData("NA", map.get("RuleType"), 17);  //'Actual' column=17
            ExcelUtils.writeExcelData("RuleName not found", map.get("RuleType"), 18); //'Pass/Fail' column=18
            System.out.println("Fail");
        }
    }

    @AfterMethod
    public void closeBrowser(){
        getDriver().close();
    }

}
