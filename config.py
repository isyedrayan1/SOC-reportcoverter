# Slide configuration for the PPTX report
SLIDE_CONFIG = [
    {"slide_number": 1, "title": "", "type": "title", "data_key": "month_year"},
    {"slide_number": 3, "title": "Alert Classification", "type": "bar_chart", "data_key": "priority_breakdown", "chart_title": "Alerts by Priority", "categories": ["Low", "Medium", "High"]},
    {"slide_number": 4, "title": "Payrollorg SOC Report: Alerts Comparison", "type": "bar_chart", "data_key": "alerts_comparison", "chart_title": "Comparison of Alerts Count from Past 2 Months", "categories": ["March 1st - March 31st", "April 1st - April 30th"]},
    {"slide_number": 5, "title": "Payrollorg SOC Report: Alerts Escalated Comparison", "type": "bar_chart", "data_key": "escalated_comparison", "chart_title": "Escalated Alerts over the Past 2 Months", "categories": ["March 1st - March 31st", "April 1st - April 30th"]},
    {"slide_number": 6, "title": "Closed Alerts by Disposition", "type": "bar_chart", "data_key": "disposition_breakdown", "chart_title": "Disposition Breakdown", "categories": ["Benign", "False Positive", "Reported Malicious"]},
    {"slide_number": 7, "title": "Projecting", "type": "static_text", "content": "Projection Data TBD"},
    {"slide_number": 8, "title": "Event Sources", "type": "table", "data_key": "event_sources", "columns": ["Event Source", "Collection", "Status"]},
    {"slide_number": 9, "title": "Attack Trends", "type": "bar_chart", "data_key": "top_alerts", "chart_title": "Top Alert Types", "categories": "dynamic"},
    {"slide_number": 10, "title": "SOC Monthly Dashboards Status", "type": "static_text", "content": "Dashboard Status TBD"},
    {"slide_number": 11, "title": "Accounts Created/Locked", "type": "text", "data_key": "accounts", "format": "Accounts Created: {created}\nAccounts Locked: {locked}"},
    {"slide_number": 12, "title": "Placeholder Slide 12", "type": "static_text", "content": "TBD"},
    {"slide_number": 13, "title": "Placeholder Slide 13", "type": "static_text", "content": "TBD"},
    {"slide_number": 14, "title": "Placeholder Slide 14", "type": "static_text", "content": "TBD"},
    {"slide_number": 15, "title": "Payrollorg Password Reset", "type": "text", "data_key": "password_resets", "format": "Password Resets: {value}"},
    {"slide_number": 16, "title": "Payrollorg Members Added to Security Groups", "type": "text", "data_key": "admin_users", "format": "Members Added: {value}"},
    {"slide_number": 17, "title": "Payrollorg Password Set Never to Expire", "type": "text", "data_key": "non_expiring", "format": "Accounts: {value}"},
    {"slide_number": 18, "title": "Payrollorg Remote Access and Cloud Storage Solutions", "type": "static_text", "content": "Remote Access Incidents: 1500"},
    {"slide_number": 19, "title": "Rapid Updates ({month})", "type": "static_text", "content": "New ABA detection rules (10 threats)\nImproved Microsoft/Office 365 event sources\nSettings UI updates"},
    {"slide_number": 20, "title": "Payrollorg Q and A", "type": "static_text", "content": "10\n\n20"},
    {"slide_number": 21, "title": "You", "type": "static_text", "content": "You"}
]