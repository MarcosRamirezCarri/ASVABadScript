import requests
from datetime import datetime
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

# Your Memberful endpoint and headers
url = 'https://asvabadvantage.memberful.com/api/graphql'
headers = {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer MGKVnam26x6dapNgwi48YKwX',
}

# Template for the GraphQL query with pagination
query_template = """
{
  members(first: 100, after: "%s") {
    edges {
      node {
        email
        trackingParams
        totalSpendCents
        subscriptions {
          activatedAt
          orders {
            createdAt
            status
            totalCents
            type
          }
        }
      }
    }
    pageInfo {
      endCursor
      hasNextPage
    }
  }
}
"""

# Initialize pagination variables
has_next_page = True
end_cursor = None
members_by_month = defaultdict(list)
all_emails = set()
all_months = set()
total_members = 0
paid_members = set()
members_2023_onwards = set()

# Fetch all members using pagination
while has_next_page:
    # Update the query with the current end_cursor
    query = {"query": query_template % (end_cursor if end_cursor else "")}
    response = requests.post(url, headers=headers, json=query)
    data = response.json()
    
    # Process the current page of data
    for member in data['data']['members']['edges']:
        total_members += 1
        email = member['node']['email']
        subscription = member['node']['subscriptions'][0] if member['node']['subscriptions'] else None
        totalSpend = member['node']['totalSpendCents']
        
        if totalSpend > 0:
            paid_members.add(email)
            if subscription and subscription.get("activatedAt"):
                activated_date = datetime.fromtimestamp(int(subscription['activatedAt']))
                if activated_date.year >= 2023:
                    members_2023_onwards.add(email)
                    year_month = activated_date.strftime('%Y-%m')
                    members_by_month[year_month].append(email)
                    all_emails.add(email)
                    all_months.add(year_month)
    
    # Prepare for the next iteration
    page_info = data['data']['members']['pageInfo']
    has_next_page = page_info['hasNextPage']
    end_cursor = page_info['endCursor']

# Create Excel workbook
wb = Workbook()
ws = wb.active
ws.title = "Subscribers by Month"

# Sort months and emails
sorted_months = sorted(all_months)
sorted_emails = sorted(all_emails)

# Define styles
header_fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
header_font = Font(bold=True)
checkmark_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")

# Write month headers
for col, month in enumerate(sorted_months, start=2):
    cell = ws.cell(row=1, column=col, value=month)
    cell.fill = header_fill
    cell.font = header_font

# Write email header
ws.cell(row=1, column=1, value="Email").fill = header_fill
ws.cell(row=1, column=1, value="Email").font = header_font

# Add revenue row
ws.cell(row=2, column=1, value="Monthly Revenue ($)").font = header_font
for col, month in enumerate(sorted_months, start=2):
    subscriber_count = len(members_by_month[month])
    revenue = subscriber_count * 11
    ws.cell(row=2, column=col, value=revenue)

# Fill in the data (starting from row 3 now)
for row, email in enumerate(sorted_emails, start=3):
    ws.cell(row=row, column=1, value=email)
    for col, month in enumerate(sorted_months, start=2):
        if email in members_by_month[month]:
            cell = ws.cell(row=row, column=col, value="✓")
            cell.fill = checkmark_fill

# Adjust column widths
ws.column_dimensions[get_column_letter(1)].width = 35  # Email column
for col in range(2, len(sorted_months) + 2):
    ws.column_dimensions[get_column_letter(col)].width = 15

# Save the workbook
wb.save('subscribers_by_month.xlsx')
print("Excel file 'subscribers_by_month.xlsx' has been created successfully.")

