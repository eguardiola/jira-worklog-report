package es.bmcs.jiraworklogreport;

import java.util.ArrayList;
import java.util.List;
import net.rcarz.jiraclient.BasicCredentials;
import net.rcarz.jiraclient.Issue;
import net.rcarz.jiraclient.Issue.SearchResult;
import net.rcarz.jiraclient.JiraClient;
import net.rcarz.jiraclient.WorkLog;

public class WorkLogReportApp {
	
	public static void main(String[] args) {
	  BasicCredentials creds = new BasicCredentials(Config.UserName, Config.Password);
      JiraClient jira = new JiraClient(Config.JiraServerUri, creds);
      
      ArrayList<WorkLogDTO> worklogDtos = new ArrayList<WorkLogDTO>();
      
      try {
        SearchResult searchResult = jira.searchIssues(Config.jql);
        for (Issue issue : searchResult.issues) {
          List<WorkLog> worklogs = issue.getAllWorkLogs();
//          if (issue.getWorkLogs().size() == 20) {
//            worklogs = issue.getAllWorkLogs();
//          }
          
          for (WorkLog worklog : worklogs) {
            
            WorkLogDTO worklogDto = new WorkLogDTO();
            worklogDto.setProjectName(issue.getProject().getName());
            worklogDto.setIssueKey(issue.getKey());
            worklogDto.setIssueType(issue.getIssueType().getName());
            worklogDto.setComponent(issue.getComponents().iterator().next().getName());
            //worklogDto.setDepartment(issue.getField("customfield_10100").toString());
            worklogDto.setAuthor(worklog.getAuthor().getDisplayName());
            worklogDto.setStartDate(worklog.getUpdatedDate());
            worklogDto.setYear(worklog.getUpdatedDate().getYear() + 1900);
            worklogDto.setMonth(worklog.getUpdatedDate().getMonth() + 1);
            worklogDto.setMinutesSpent(worklog.getTimeSpentSeconds() / 60);
            
            worklogDtos.add(worklogDto);
          }
        }
      } catch (Exception e) {
        e.printStackTrace();
      }
      finally {
        System.out.println("Generating Excel report...");
        ExcelReporter reporter = new ExcelReporter("worklogReport.xls");
        try {
            reporter.writeReportToExcel(worklogDtos);
            reporter.closeWorksheet();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
          System.out.println("All done.");
        }
      }
	}

//	private static String getDepartment(Issue issue) {
//		Field field = issue.getFieldByName("Departamento");
//		if (field != null) {
//			JSONObject jsonObject = (JSONObject) field.getValue();
//			String value = null;
//			try {
//				value = jsonObject.get("value").toString();
//			} catch (JSONException e) {
//			}
//			return value;
//		}
//		
//		return null;
//	}
}
