package es.bmcs.jiraworklogreport;

import java.util.ArrayList;
import java.util.concurrent.ExecutionException;
import org.codehaus.jettison.json.JSONException;
import org.codehaus.jettison.json.JSONObject;
import com.atlassian.jira.rest.client.api.JiraRestClient;
import com.atlassian.jira.rest.client.api.domain.Issue;
import com.atlassian.jira.rest.client.api.domain.IssueField;
import com.atlassian.jira.rest.client.internal.async.AsynchronousJiraRestClientFactory;

public class WorkLogReportApp {

  public static void main(String[] args) throws InterruptedException, ExecutionException {
    final AsynchronousJiraRestClientFactory factory = new AsynchronousJiraRestClientFactory();
    final JiraRestClient restClient = factory
        .createWithBasicHttpAuthentication(Config.JiraServerUri, Config.UserName, Config.Password);

    ArrayList<WorkLog> worklogs = new ArrayList<WorkLog>();

    restClient.getSearchClient()
        .searchJql("worklogAuthor in ('eguardiola','hanhelsing','pamp6675','fguerrerom')", 1000, 0,
            null)
        .get().getIssues().forEach(bi -> {
          try {

            Issue issue = restClient.getIssueClient().getIssue(bi.getKey()).get();

            issue.getWorklogs().forEach(wl -> {
              WorkLog workLog = new WorkLog();
              workLog.setProjectName(issue.getProject().getName());
              workLog.setIssueKey(issue.getKey());
              workLog.setIssueType(issue.getIssueType().getName());
              workLog.setComponent(issue.getComponents().iterator().next().getName());
              workLog.setDepartment(getDepartment(issue));
              workLog.setAuthor(wl.getAuthor().getDisplayName());
              workLog.setStartDate(wl.getStartDate());
              workLog.setYear(wl.getStartDate().getYear());
              workLog.setMonth(wl.getStartDate().getMonthOfYear());
              workLog.setMinutesSpent(wl.getMinutesSpent());
              worklogs.add(workLog);

              System.out.println(workLog);
            });
          } catch (Exception e) {
            e.printStackTrace();
          }
        });

    System.out.println("Generating Excel report...");
    ExcelReporter reporter = new ExcelReporter("worklogReport.xls");
    try {
      reporter.writeReportToExcel(worklogs);
      reporter.closeWorksheet();
    } catch (Exception e) {
      e.printStackTrace();
    }

    System.out.println("All done.");
  }

  private static String getDepartment(Issue issue) {
    IssueField field = issue.getFieldByName("Departamento");
    if (field != null) {
      JSONObject jsonObject = (JSONObject) field.getValue();
      String value = null;
      try {
        value = jsonObject.get("value").toString();
      } catch (JSONException e) {
      }
      return value;
    }

    return null;
  }
}
