/**********************************************************
@class      AbstractBaseBatch

@brief      AbstractBaseBatch Class 
Used as a base class to automatically capture Execution Logs during start, execute, Finish methods.

@Test        
**********************************************************/
public abstract class AbstractBaseBatch implements Database.Batchable<sObject>, Database.Stateful, Database.AllowsCallouts, Database.RaisesPlatformEvents {
    public LogEvent__c batchLog ;
    public String customUUID;
    //Constructor to get CustomUUID Value;
    public AbstractBaseBatch(String uuid){
        this.customUUID = uuid;
    }
    // Default Constructor
    public AbstractBaseBatch(){
       
    }
    /****************************************************************
start method 
paramters : batchableContext
*****************************************************************/
    public Database.QueryLocator start(Database.BatchableContext bc) {
        // Create a log record and capture it in the class-level variable
        String className = String.valueOf(this).substring(0,String.valueOf(this).indexOf(':'));
        batchLog = ExecutionLogger.onBatchJobStart(bc,customUUID,className);
        return startBatch(bc);
    }
    
    /****************************************************************
execute method 
paramters : batchableContext, Scope
*****************************************************************/
    public void execute(Database.BatchableContext bc, List<sObject> scope) {
        try {
            executeBatch(bc, scope);
        } catch (Exception ex) {throw ex;}
    }
    
    /****************************************************************
finish method 
paramters : batchableContext
*****************************************************************/
    public void finish(Database.BatchableContext bc) {
        // Read details from AsyncApexJobs and update the same details in the log object for the batch
        ExecutionLogger.onBatchJobFinish(bc, batchLog); 
        finishBatch(bc);
    }
    
    // Abstract method - startBatch. 
    // This method needs to be override in batch classes which are extending this abstarct class
    public abstract Database.QueryLocator startBatch(Database.BatchableContext bc);
    
    // Abstract method - executeBatch
    // This method needs to be override in batch classes which are extending this abstarct class
    public abstract void executeBatch(Database.BatchableContext bc, List<sObject> scope);
    
    // Abstract method - finishBatch
    // This method needs to be override in batch classes which are extending this abstarct class
    public abstract void finishBatch(Database.BatchableContext bc);
}