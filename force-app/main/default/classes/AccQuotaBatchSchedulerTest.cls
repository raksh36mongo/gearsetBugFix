/**
 * @group test
 * @description test class for AccountQuotaBatchScheduler
 * @see AccountQuotaBatchScheduler
 */
@isTest
public without sharing class AccQuotaBatchSchedulerTest {
  public static final String CRON_EXP = '0 25 * * * ?';
  /**
   * @description method to test scheduler
   */
  @isTest
  public static void testScheduledJob() {
    AccountQuotaBatchScheduler sync = new AccountQuotaBatchScheduler();
    Test.startTest();
    String jobId = System.schedule('Sync Updates', CRON_EXP, sync);
    Test.stopTest();
    System.assertEquals(false, String.isBlank(jobId), 'Job not scheduled');
  }
}