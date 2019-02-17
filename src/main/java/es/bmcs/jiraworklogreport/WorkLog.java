package es.bmcs.jiraworklogreport;

import org.apache.commons.lang.builder.ToStringBuilder;
import org.joda.time.DateTime;

public class WorkLog {
	private String ProjectName;
	private String IssueKey;
	private String IssueType;
	private String Component;
	private String Department;
	private String Author;
	private DateTime StartDate;
	private int MinutesSpent;
	private int Year;
	private int Month;

	public String getProjectName() {
		return ProjectName;
	}

	public void setProjectName(String projectName) {
		ProjectName = projectName;
	}

	public int getMinutesSpent() {
		return MinutesSpent;
	}

	public void setMinutesSpent(int i) {
		MinutesSpent = i;
	}

	public DateTime getStartDate() {
		return StartDate;
	}

	public void setStartDate(DateTime dateTime) {
		StartDate = dateTime;
	}

	public String getAuthor() {
		return Author;
	}

	public void setAuthor(String author) {
		Author = author;
	}

	public String getDepartment() {
		return Department;
	}

	public void setDepartment(String department) {
		Department = department;
	}

	public String getComponent() {
		return Component;
	}

	public void setComponent(String component) {
		Component = component;
	}

	public String getIssueType() {
		return IssueType;
	}

	public void setIssueType(String issueType) {
		IssueType = issueType;
	}

	public String getIssueKey() {
		return IssueKey;
	}

	public void setIssueKey(String issueKey) {
		IssueKey = issueKey;
	}

	@Override
	public String toString() {
		return ToStringBuilder.reflectionToString(this);

	}

	public int getYear() {
		return Year;
	}

	public void setYear(int year) {
		Year = year;
	}

	public int getMonth() {
		return Month;
	}

	public void setMonth(int month) {
		Month = month;
	}


}