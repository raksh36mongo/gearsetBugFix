@isTest
public class AccountDAndBCheckBatchTest {
    
    @testSetup static void createTestData() {
        ByPassValidationRules__c bypass = new ByPassValidationRules__c();
        bypass.SetupOwnerId = UserInfo.getUserId();
        bypass.Opportunity__c = 'X79;X19';
        bypass.SBQQ_QuoteLine__c = 'X04';
        insert bypass;
    }
    public static void createCustomSettingsData() {
        Org_Settings__c orgSetting = new Org_Settings__c(
                                        D_B_Integrations_Active__c = TRUE,
                                        D_B_DirectPlus_Authentication_Endpoint__c = 'https://plus.dnb.com/v2/token',
                                        D_B_DirectPlus_CMPTCS_Endpoint__c = 'https://plus.dnb.com/v1/data/duns/{dunsNumber}?productId=cmptcs&versionId=v1&tradeUp=hq&orderReason=6332',
                                        D_B_DirectPlus_Consumer_Key__c = 'ConsumerKey',
                                        D_B_DirectPlus_Consumer_Secret__c = 'Secret',
                                        D_B_DirectPlus_Match_Endpoint__c = 'https://plus.dnb.com/v1/match/cleanseMatch');
        insert orgSetting;

        //Create Custom Settings for Country Coverage Models
        CountryCS__c ccs1 = new CountryCS__c(Name = 'USA', Price_Book_ID__c = test.getStandardPricebookId(), ISO3__c = 'USA', ISO2__c = 'US', Geo_Region__c = 'Americas', Geo_Area__c = 'United States', Country_Name__c = 'United States', Account_Coverage_Model__c = 'CorpAuto', Enterprise_Threshold__c = 750000000, Enterprise_Threshold_Employees__c = 3750);
        CountryCS__c ccs2 = new CountryCS__c(Name = 'GBR', Price_Book_ID__c = test.getStandardPricebookId(), ISO3__c = 'GBR', ISO2__c = 'UK', Geo_Region__c = 'EMEA', Geo_Area__c = 'NEUR', Country_Name__c = 'United Kingdom', Account_Coverage_Model__c = 'EntUnassigned', Enterprise_Threshold__c = 500000000, Enterprise_Threshold_Employees__c = 2500);
        upsert new List<CountryCS__c>{ccs1, ccs2} Name;
    }
    
    @isTest
    static void testBatch() {
        DAndBUtility.isRunningDnBTest = TRUE;
        Account acc = new Account();
        acc.Name = 'Account 1';
        acc.Account_Street__c = '81 newgate street';
        acc.Account_City__c = 'London';
        acc.BillingCountry = 'GBR';
        acc.Account_State__c = 'London';
        acc.Account_ZIP_Code__c = '12345';
        acc.DUNS_Number__c = '398173534';
        acc.Projected_Atlas_Usage__c= 27000;
        acc.credit_status__c= null;
        acc.D_B_DUNS_Retrieve_Date__c = null;
        acc.D_B_Profile_Retrieve_Date__c = null;
        acc.Credit_Review_Date__c = Date.newInstance(2023,09,08);
        insert acc;
        Test.startTest();
        AccountDAndBCheckBatch batchScheduler = new AccountDAndBCheckBatch();
        System.schedule('Test Batch Scheduler Job for D&B', '0 0 12 * * ?', batchScheduler);
        Test.stopTest();
    }
}