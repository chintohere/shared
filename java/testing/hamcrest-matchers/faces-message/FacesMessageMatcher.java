package au.com.hrx.search_string_builder;

import javax.faces.application.FacesMessage;
import javax.faces.application.FacesMessage.Severity;

import org.hamcrest.Description;
import org.hamcrest.TypeSafeDiagnosingMatcher;

public class FacesMessageMatcher extends TypeSafeDiagnosingMatcher<FacesMessage>
{
	String summary;

	Severity severity;

	public FacesMessageMatcher(Severity severity, String summary)
	{
		super(FacesMessage.class);
		this.summary = summary;
		this.severity = severity;
	}

	@Override
	protected boolean matchesSafely(FacesMessage item, Description mismatchDescription)
	{
		boolean severityMatches = item.getSeverity().equals(severity);
		if (!severityMatches) mismatchDescription.appendText(item.getSeverity() + " vs " + severity);
		boolean summaryMatches = item.getSummary().equals(summary);
		if (!summaryMatches) mismatchDescription.appendText(item.getSummary() + " vs " + summary);

		return summaryMatches && severityMatches;
	}

	@Override
	public void describeTo(Description description)
	{
		description.appendText(summary);
	}

}

