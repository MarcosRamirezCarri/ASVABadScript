# ASVAB Subscriber Tracking Script

This Python script tracks and analyzes subscriber data from the ASVAB Advantage Memberful platform. It generates a detailed Excel report showing subscriber activity and revenue over time, tracking all subscriptions regardless of when they started.

## Quick Start

1. Set your Bearer token in the script:
   ```python
   BEARER_TOKEN = 'your-token-here'  # in asvab.py
   ```
   Or better yet, use environment variables:
   ```bash
   export ASVAB_BEARER_TOKEN='your-token-here'
   ```

2. Run the script:
   ```bash
   python asvab.py
   ```

## Features

- Connects to the ASVAB Advantage Memberful API using GraphQL
- Tracks ALL subscriber data including:
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
- Monthly revenue calculations ($11 per subscriber)
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

The script uses a Bearer token for API authentication. You can set it in two ways:
1. Directly in the script: `BEARER_TOKEN = 'your-token-here'`
2. Using environment variables: `export ASVAB_BEARER_TOKEN='your-token-here'`

## Data Processing

The script:
1. Fetches all members using pagination
2. Tracks all paid members and their subscription dates
3. Records both initial subscriptions and renewals
4. Organizes data by month
5. Calculates monthly revenue ($11 per active subscriber)
6. Generates a formatted Excel report

## Excel Report Format

- Column A: Email addresses
- Subsequent columns: Months (YYYY-MM format)
- Row 2: Monthly revenue calculations
- Row 3: Number of subscribers per month
- Checkmarks (✓) indicate active subscriptions
- Gray formatting for headers and checkmarks
- Auto-sized columns for better readability

## Inclusion Criteria

A subscriber will appear in the Excel report if they meet ALL of the following conditions:

1. **Payment Status**:
   - Must have made at least one payment (totalSpend > 0)

2. **Subscription Data**:
   - Must have a valid subscription with an activation date
   - Must have valid order records with creation dates

3. **Monthly Tracking**:
   - A checkmark (✓) appears in months where:
     - Initial subscription was activated
     - Continued subscription orders were placed
     - Different months are tracked separately (e.g., if someone subscribed in Jan 2023 and renewed in Feb 2023, they'll have checkmarks in both months)
