# ASVAB Subscriber Tracking Script

This Python script tracks and analyzes subscriber data from the ASVAB Advantage Memberful platform. It generates a detailed Excel report showing subscriber activity and revenue over time.

## Features

- Connects to the ASVAB Advantage Memberful API using GraphQL
- Tracks subscriber data including:
  - Email addresses
  - Subscription activation dates
  - Total spend
  - Order history
- Generates an Excel report with:
  - Monthly subscriber tracking
  - Revenue calculations
  - Visual checkmarks for active subscriptions
  - Formatted headers and columns

## Output

The script generates an Excel file (`subscribers_by_month.xlsx`) that contains:
- A matrix of subscribers by month
- Monthly revenue calculations
- Visual indicators (checkmarks) for active subscriptions
- Formatted headers and properly sized columns

## Requirements

- Python 3.x
- Required Python packages:
  - requests
  - openpyxl
  - datetime
  - collections

## Authentication

The script uses a Bearer token for API authentication. The token is currently hardcoded in the script.

## Data Processing

The script:
1. Fetches all members using pagination
2. Tracks paid members and their subscription dates
3. Organizes data by month
4. Calculates monthly revenue
5. Generates a formatted Excel report

## Excel Report Format

- Column A: Email addresses
- Subsequent columns: Months (YYYY-MM format)
- Row 2: Monthly revenue calculations
- Checkmarks (✓) indicate active subscriptions
- Gray formatting for headers and checkmarks
- Auto-sized columns for better readability
