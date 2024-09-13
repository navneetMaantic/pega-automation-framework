package com.maantic.automation.tests;

import com.maantic.automation.base.BasePage;
import com.maantic.automation.listeners.AllureListeners;
import com.maantic.automation.pages.DevStudioPage;
import com.maantic.automation.pages.LoginPage;
import com.maantic.automation.pages.SLAPage;
import com.maantic.automation.utils.Constants;
import com.maantic.automation.utils.ExcelDataProvider;
import com.maantic.automation.utils.ExcelUtils;
import io.qameta.allure.Description;
import io.qameta.allure.Severity;
import io.qameta.allure.SeverityLevel;
import org.testng.Assert;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Listeners({ AllureListeners.class })
public class GBTSLATest2 extends BasePage {

	LoginPage loginPage = new LoginPage();
	DevStudioPage homePage = new DevStudioPage();
	SLAPage slaPage = new SLAPage();
	String testUsername, testPassword;

	@Test(priority = 1, groups = { "Sanity",
			"Regression" }, description = "fetch URL", dataProviderClass = ExcelDataProvider.class, dataProvider = "ExcelTestDataLogin")
	public void getURL(Map<String, String> map) throws InterruptedException {
		System.out.println("Test starts");
		appUrl = map.get("URL");
		testUsername = map.get("Username");
		testPassword = map.get("Password");
	}

	@Test(priority = 2)
	@Severity(SeverityLevel.NORMAL)
	@Description("Test Case Description: Verify the SLA is working properly.")
	public void SLA() throws InterruptedException {
		List<Map<String, String>> list = ExcelUtils.getExcelData(Constants.EXCEL_SHEET_NAME);
		List<Map<String, String>> smallList = new ArrayList<>();
		Map<String, String> map = null;
		int j = 0, count = 0;
		for (int i2 = 0; i2 < list.size(); i2++) {
			if (list.get(i2).get("RuleType").trim().equalsIgnoreCase("SLA")) {
				System.out.println("Count " + count);
				smallList.add(list.get(i2));
				map = smallList.get(j);
				j++;
				// ************************************************//
				if (count == 0) {
					loginPage.enterUserNameText(testUsername);
					loginPage.enterPasswordText(testPassword);
					loginPage.clickOnLogInButton();
					Thread.sleep(5000);
				}

				// if correct ruleName present in file
				if (homePage.selectRuleType(count, map.get("RuleType"), map.get("RuleName"),
						map.get("RuleSetVersion"))) {
					String sla_output = "true";
					for (int i = 1; i <= 12; i++) {
						if (!slaPage.validateSLAValues(i, map.get("Param" + i + ""))) {
							sla_output = "false";
							slaPage.closeSLAPage();
							break;
						}
					}
					slaPage.closeSLAPage();
					// writing 'actual' result in file
					Thread.sleep(3000);
					ExcelUtils.writeExcelData(sla_output, map.get("RuleType"), 17); // 'Actual' column=17
					Thread.sleep(2000);
					// writing 'pass/fail' in file
					if (map.get("Expected").equalsIgnoreCase(sla_output)) {
						ExcelUtils.writeExcelData("Pass", map.get("RuleType"), 18); // 'Pass/Fail' column=18
						System.out.println("Pass");
					} else {
						ExcelUtils.writeExcelData("Fail", map.get("RuleType"), 18); // 'Pass/Fail' column=18
						System.out.println("Fail");
					}
				} else {// if incorrect ruleName present in file
					ExcelUtils.writeExcelData("NA", map.get("RuleType"), 17); // 'Actual' column=17
					ExcelUtils.writeExcelData("RuleName not found", map.get("RuleType"), 18); // 'Pass/Fail' column=18
					System.out.println("Fail");
				}
				count++;
			}
		}
	}

	@AfterMethod
	public void closeBrowser() {
		getDriver().close();
	}

}
