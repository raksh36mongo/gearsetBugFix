/*******************************************************//**
@class  AccountDAndBCheckBatch
This class will be called from AccountDAndBCheckScheduler   
Changelog
2025-03-11  Nikhil Palekar   CRMENG-21477
***********************************************************/

public class AccountDAndBCheckBatch implements Database.Batchable<sObject>, Database.AllowsCallouts, Schedulable {

    /**
     * To Schedule the Batch Job
     * @param sc SchedulableContext
     */
    public void execute(SchedulableContext sc) {
        Integer customBatchSize= Integer.valueOf(System.label.QTC_Batch_Size_for_New_D_And_B_Logic);
        Integer flexQueueCount = [SELECT COUNT() FROM AsyncApexJob WHERE JobType = 'BatchApex' AND Status IN ('Holding', 'Queued')];
            
        if (flexQueueCount >= 100) {
            // Flex queue has 100 or more jobs. Scheduling batch to run in 15 minutes.
            System.scheduleBatch(new AccountDAndBCheckBatch(), 'AccountDAndBCheckBatch', 15, customBatchSize);
        } else {
            // Flex queue has capacity. Executing batch now.
            Database.executeBatch(new AccountDAndBCheckBatch(), customBatchSize);
        }
    }
    
    /**
     * Prepares the batch job by providing the scope of records
     * @param context BatchableContext
     */
    public Database.QueryLocator start(Database.BatchableContext context) {
        Date dateTenMonthAgo= System.today().addMonths(-10);
        return Database.getQueryLocator([SELECT Id, Credit_Status__c, Credit_Review_Date__c, Projected_Atlas_Usage__c FROM Account WHERE (Projected_Atlas_Usage__c >= 25000 AND Credit_Review_Date__c < :dateTenMonthAgo) OR (Projected_Atlas_Usage__c >= 25000 AND credit_status__c = null)]);
    }

    /**
     * Processes each batch of accounts, with a batch size of 40
     * @param scope List of Account records
     * @param context BatchableContext
     */
    public void execute(Database.BatchableContext context, List<Account> accountsToProcess){
        Set<Id> accountIds = new Set<Id>();
        for(Account acc : accountsToProcess){
            accountIds.add(acc.Id);
        }
        // Process D And B check for Accounts
        DAndBAccountsProcessing.processAccounts(accountIds, 'AccountDAndBCheckBatch');
    }

    /**
     * Finalizes the batch job and logs completion
     * @param context BatchableContext
     */
    public void finish(Database.BatchableContext context) {
    }
}