import os
import requests
import pandas as pd
import time
import logging
from datetime import datetime
from dateutil.relativedelta import relativedelta
import json
import sys
from dotenv import load_dotenv

# Set up logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('debug.log'),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger('github_analyzer')

class GitHubContributionAnalyzer:
    def __init__(self, token, organizations, time_period_months=None):
        """
        Initialize the analyzer with GitHub token and organizations.
        
        Args:
            token (str): GitHub personal access token
            organizations (list): List of GitHub organization names
            time_period_months (int, optional): Number of months to look back for contributions
        """
        self.headers = {'Authorization': f'token {token}'}
        self.organizations = organizations if isinstance(organizations, list) else [organizations]
        self.base_url = "https://api.github.com"
        self.rate_limit_remaining = 5000  # Default GitHub API rate limit
        self.rate_limit_reset = 0
        
        logger.info(f"Initializing analyzer for organizations: {', '.join(self.organizations)}")
        self.update_rate_limit_info()
        
        # Set time period filter if specified
        self.time_filter = ""
        if time_period_months:
            since_date = (datetime.now() - relativedelta(months=time_period_months)).strftime('%Y-%m-%d')
            self.time_filter = f"+created:>={since_date}"
            logger.info(f"Setting time filter to {time_period_months} months (since {since_date})")
            
        # Cache for API responses to avoid duplicate requests
        self.cache = {}
        
        # For saving progress
        self.completed_users = set()
        self.progress_file = "github_analysis_progress.json"
        self.load_progress()
        
        # Current date/time for reporting
        self.analysis_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        logger.info(f"Analysis started at: {self.analysis_date}")

    def update_rate_limit_info(self):
        """Update information about the current rate limit status"""
        try:
            response = requests.get(f"{self.base_url}/rate_limit", headers=self.headers)
            if response.status_code == 200:
                data = response.json()
                self.rate_limit_remaining = data['resources']['core']['remaining']
                self.rate_limit_reset = data['resources']['core']['reset']
                logger.info(f"API Rate Limit: {self.rate_limit_remaining} requests remaining, resets at {datetime.fromtimestamp(self.rate_limit_reset).strftime('%Y-%m-%d %H:%M:%S')}")
            else:
                logger.warning(f"Couldn't fetch rate limit info. Status code: {response.status_code}")
        except Exception as e:
            logger.error(f"Error checking rate limits: {str(e)}")

    def wait_for_rate_limit(self):
        """Wait if we're close to hitting the rate limit"""
        if self.rate_limit_remaining < 10:
            current_time = time.time()
            if current_time < self.rate_limit_reset:
                wait_time = self.rate_limit_reset - current_time + 5  # Add 5 seconds buffer
                logger.info(f"Rate limit almost reached. Waiting for {wait_time:.1f} seconds until reset...")
                time.sleep(wait_time)
                self.update_rate_limit_info()
                logger.info("Resuming API calls...")
    
    def save_progress(self):
        """Save the current progress to a file"""
        progress_data = {
            'completed_users': list(self.completed_users),
            'last_updated': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        try:
            with open(self.progress_file, 'w') as f:
                json.dump(progress_data, f)
            logger.debug(f"Progress saved: {len(self.completed_users)} users completed")
        except Exception as e:
            logger.error(f"Failed to save progress: {str(e)}")
            
    def load_progress(self):
        """Load progress from file if it exists"""
        try:
            if os.path.exists(self.progress_file):
                with open(self.progress_file, 'r') as f:
                    progress_data = json.load(f)
                    self.completed_users = set(progress_data.get('completed_users', []))
                    last_updated = progress_data.get('last_updated', 'unknown')
                    if self.completed_users:
                        logger.info(f"Loaded progress: {len(self.completed_users)} users already analyzed (last updated: {last_updated})")
        except Exception as e:
            logger.warning(f"Couldn't load progress file: {str(e)}")
    
    def get_user_stats(self, username):
        """
        Get comprehensive statistics for a specific user across multiple organizations.
        
        Args:
            username (str): GitHub username
            
        Returns:
            dict: Dictionary containing user statistics
        """
        stats = {
            'username': username,
            'pr_total': 0,
            'pr_merged': 0,
            'pr_open': 0,
            'commits': 0,
            'issues_opened': 0,
            'issues_closed': 0,
            'issues_commented': 0,
            'repos_contributed': set(),
            'reviews_submitted': 0,
            'organizations': {}
        }
        
        # Track metrics separately for each organization
        for org in self.organizations:
            stats['organizations'][org] = {
                'pr_total': 0,
                'pr_merged': 0,
                'pr_open': 0,
                'commits': 0,
                'issues_opened': 0,
                'issues_closed': 0,
                'issues_commented': 0,
                'repos_contributed': set(),
                'reviews_submitted': 0
            }
        
        try:
            for org in self.organizations:
                logger.info(f"Analyzing {username}'s contributions to {org}...")
                org_stats = stats['organizations'][org]
                
                # Fetch PRs created by user
                prs_created_url = f"{self.base_url}/search/issues?q=author:{username}+org:{org}+is:pr{self.time_filter}&per_page=100"
                pr_data = self._paginated_request(prs_created_url)
                
                if pr_data and 'items' in pr_data:
                    pr_items = pr_data['items']
                    org_stats['pr_total'] = len(pr_items)
                    stats['pr_total'] += org_stats['pr_total']
                    
                    # Process each PR to get more details
                    for pr in pr_items:
                        if pr.get('state') == 'open':
                            org_stats['pr_open'] += 1
                            stats['pr_open'] += 1
                        elif pr.get('state') == 'closed':
                            # Check if PR was merged using pull_request.merged_at field
                            if pr.get('pull_request', {}).get('merged_at'):
                                org_stats['pr_merged'] += 1
                                stats['pr_merged'] += 1
                            
                        # Extract repository name from PR URL
                        repo_url = pr.get('repository_url', '')
                        if repo_url:
                            repo_name = repo_url.split('/')[-1]
                            org_stats['repos_contributed'].add(repo_name)
                            stats['repos_contributed'].add(f"{org}/{repo_name}")
                
                # Fetch issues created by user
                issues_url = f"{self.base_url}/search/issues?q=author:{username}+org:{org}+is:issue{self.time_filter}&per_page=100"
                issues_data = self._paginated_request(issues_url)
                
                if issues_data and 'items' in issues_data:
                    issue_items = issues_data['items']
                    org_stats['issues_opened'] = len(issue_items)
                    stats['issues_opened'] += org_stats['issues_opened']
                    
                    # Count closed issues
                    for issue in issue_items:
                        if issue.get('state') == 'closed':
                            org_stats['issues_closed'] += 1
                            stats['issues_closed'] += 1
                            
                        # Extract repository info
                        repo_url = issue.get('repository_url', '')
                        if repo_url:
                            repo_name = repo_url.split('/')[-1]
                            org_stats['repos_contributed'].add(repo_name)
                            stats['repos_contributed'].add(f"{org}/{repo_name}")
                
                # Get commit count
                # Note: This is approximate as GitHub API doesn't provide a direct way to count all commits across an org
                repo_batch_size = 3  # Limit batch size to manage rate limits
                repo_list = list(org_stats['repos_contributed'])
                
                for i in range(0, len(repo_list), repo_batch_size):
                    batch_repos = repo_list[i:i+repo_batch_size]
                    for repo_name in batch_repos:
                        commits_url = f"{self.base_url}/repos/{org}/{repo_name}/commits?author={username}&per_page=100"
                        repo_commits = self._paginated_request(commits_url)
                        if isinstance(repo_commits, list):
                            org_stats['commits'] += len(repo_commits)
                            stats['commits'] += len(repo_commits)
                    
                    # Check rate limit after each batch
                    self.update_rate_limit_info()
                    self.wait_for_rate_limit()
                
                # Get review activity
                reviews_url = f"{self.base_url}/search/issues?q=reviewed-by:{username}+org:{org}+is:pr{self.time_filter}&per_page=100"
                review_data = self._paginated_request(reviews_url)
                if review_data and 'items' in review_data:
                    org_stats['reviews_submitted'] = len(review_data['items'])
                    stats['reviews_submitted'] += org_stats['reviews_submitted']
                
                # Get issue comments (participation)
                # Split into two queries to avoid API limitations
                comments_url_issues = f"{self.base_url}/search/issues?q=commenter:{username}+org:{org}+is:issue{self.time_filter}&per_page=100"
                comments_data_issues = self._paginated_request(comments_url_issues)
                
                self.wait_for_rate_limit()  # Check rate limit between queries
                
                comments_url_prs = f"{self.base_url}/search/issues?q=commenter:{username}+org:{org}+is:pr{self.time_filter}&per_page=100"
                comments_data_prs = self._paginated_request(comments_url_prs)
                
                comment_count = 0
                if comments_data_issues and 'items' in comments_data_issues:
                    comment_count += len(comments_data_issues['items'])
                if comments_data_prs and 'items' in comments_data_prs:
                    comment_count += len(comments_data_prs['items'])
                    
                org_stats['issues_commented'] = comment_count
                stats['issues_commented'] += comment_count
                
                # Convert org's repo set to count for Excel output
                org_stats['repos_contributed'] = len(org_stats['repos_contributed'])
        except Exception as e:
            logger.error(f"Error fetching stats for {username}: {str(e)}")
            raise
        
        # Convert overall repo set to count for Excel output
        stats['repos_contributed'] = len(stats['repos_contributed'])
        
        # Add organization breakdown to main stats
        for org in self.organizations:
            prefix = f"{org.replace('-', '_')}_"
            for key, value in stats['organizations'][org].items():
                stats[f"{prefix}{key}"] = value
        
        # Remove the organizations dict as we've flattened it
        del stats['organizations']
        
        logger.info(f"Completed analysis for {username}: {stats['pr_merged']} merged PRs, {stats['commits']} commits across {stats['repos_contributed']} repositories")
        return stats
    
    def analyze_users(self, usernames):
        """
        Analyze multiple users and return their statistics.
        
        Args:
            usernames (list): List of GitHub usernames
            
        Returns:
            list: List of dictionaries with user statistics
        """
        results = []
        # Filter out users that are already analyzed
        pending_users = [u for u in usernames if u not in self.completed_users]
        
        if len(pending_users) < len(usernames):
            logger.info(f"Skipping {len(usernames) - len(pending_users)} already analyzed users")
        
        for i, username in enumerate(pending_users):
            logger.info(f"[{i+1}/{len(pending_users)}] Analyzing contributions for {username}...")
            try:
                # Check if we're nearing rate limit
                self.update_rate_limit_info()
                self.wait_for_rate_limit()
                
                user_stats = self.get_user_stats(username)
                results.append(user_stats)
                
                # Mark user as completed and save progress
                self.completed_users.add(username)
                self.save_progress()
                
                logger.info(f"Successfully analyzed {username}")
            except Exception as e:
                logger.error(f"Error analyzing {username}: {str(e)}")
                
                # If we hit rate limit
                if "API rate limit exceeded" in str(e):
                    logger.warning("Rate limit exceeded. Waiting for reset...")
                    self.update_rate_limit_info()
                    self.wait_for_rate_limit()
                    
                    # Try again
                    try:
                        logger.info(f"Retrying analysis for {username}...")
                        user_stats = self.get_user_stats(username)
                        results.append(user_stats)
                        self.completed_users.add(username)
                        self.save_progress()
                        logger.info(f"Successfully analyzed {username} on retry")
                    except Exception as retry_error:
                        logger.error(f"Failed retry for {username}: {str(retry_error)}")
                        # Save a partial progress marker that we attempted this user
                        self.save_progress()
        
        # Also load results for already completed users if needed
        if self.completed_users.difference(set(pending_users)):
            logger.info("Loading cached results for previously analyzed users...")
            try:
                if os.path.exists('github_analysis_results.json'):
                    with open('github_analysis_results.json', 'r') as f:
                        cached_results = json.load(f)
                        
                    for cached_user in cached_results:
                        if cached_user['username'] in self.completed_users and cached_user['username'] not in [u['username'] for u in results]:
                            results.append(cached_user)
                            logger.debug(f"Loaded cached result for {cached_user['username']}")
            except Exception as e:
                logger.error(f"Error loading cached results: {str(e)}")
        
        # Save all results for future use
        try:
            with open('github_analysis_results.json', 'w') as f:
                json.dump(results, f)
            logger.info(f"Saved analysis results for {len(results)} users to github_analysis_results.json")
        except Exception as e:
            logger.error(f"Warning: Couldn't save results to cache file: {str(e)}")
        
        return results
    
    def export_to_excel(self, data, output_file="github_contributions.xlsx"):
        """
        Export the contribution data to an Excel file.
        
        Args:
            data (list): List of dictionaries with user statistics
            output_file (str): Path to the output Excel file
        """
        if not data:
            logger.warning("No data to export. Please check if users exist and have contributions.")
            return
            
        df = pd.DataFrame(data)
        
        # Calculate composite score (example - customize as needed)
        df['contribution_score'] = (
            (df['pr_merged'] * 3) + 
            (df['commits'] * 0.5) + 
            (df['issues_opened'] * 1) + 
            (df['issues_closed'] * 1.5) + 
            (df['reviews_submitted'] * 2) +
            (df['issues_commented'] * 0.5)
        )
        
        # Calculate org-specific scores
        for org in self.organizations:
            prefix = f"{org.replace('-', '_')}_"
            df[f'{prefix}contribution_score'] = (
                (df[f'{prefix}pr_merged'] * 3) + 
                (df[f'{prefix}commits'] * 0.5) + 
                (df[f'{prefix}issues_opened'] * 1) + 
                (df[f'{prefix}issues_closed'] * 1.5) + 
                (df[f'{prefix}reviews_submitted'] * 2) +
                (df[f'{prefix}issues_commented'] * 0.5)
            )
        
        # Sort by score
        df = df.sort_values('contribution_score', ascending=False)
        
        # Metadata for the report
        metadata = {
            'Generated On': self.analysis_date,
            'Organizations Analyzed': ', '.join(self.organizations),
            'Time Period': f"Last {self.time_filter.replace('+created:>=', '')} months" if self.time_filter else "All time",
            'Number of Contributors': len(df),
            'Tool Version': '1.2.0',
            'Generated By': 'bhavik-mangla',
        }
        
        logger.info(f"Exporting data for {len(df)} users to Excel file: {output_file}")
        
        # Create Excel writer
        try:
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                # Metadata sheet
                pd.DataFrame(list(metadata.items()), columns=['Metric', 'Value']).to_excel(
                    writer, sheet_name='Report Info', index=False
                )
                
                # Convert to Excel - main summary
                df.to_excel(writer, sheet_name='Overall Summary', index=False)
                
                # Create sheets for each organization
                for org in self.organizations:
                    prefix = f"{org.replace('-', '_')}_"
                    org_columns = ['username'] + [col for col in df.columns if col.startswith(prefix)]
                    
                    # Rename columns to remove prefix
                    org_df = df[org_columns].copy()
                    org_df.columns = [col.replace(prefix, '') if col.startswith(prefix) else col for col in org_df.columns]
                    
                    # Sort by org-specific score
                    if 'contribution_score' in org_df.columns:
                        org_df = org_df.sort_values('contribution_score', ascending=False)
                    
                    # Write to sheet
                    org_df.to_excel(writer, sheet_name=f'{org} Summary', index=False)
                
                # Access the workbook and worksheets
                workbook = writer.book
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    
                    # Add formats
                    header_format = workbook.add_format({
                        'bold': True,
                        'text_wrap': True,
                        'valign': 'top',
                        'bg_color': '#D8E4BC',
                        'border': 1
                    })
                    
                    # Get the dataframe for this sheet
                    if sheet_name == 'Report Info':
                        # Format the metadata sheet
                        worksheet.set_column('A:A', 25)
                        worksheet.set_column('B:B', 50)
                        continue
                    elif sheet_name == 'Overall Summary':
                        sheet_df = df
                    else:
                        org = sheet_name.replace(' Summary', '')
                        prefix = f"{org.replace('-', '_')}_"
                        cols = ['username'] + [col for col in df.columns if col.startswith(prefix)]
                        sheet_df = df[cols].copy()
                        sheet_df.columns = [col.replace(prefix, '') if col.startswith(prefix) else col for col in sheet_df.columns]
                    
                    # Set column width and format headers
                    for idx, col in enumerate(sheet_df.columns):
                        col_values = sheet_df[col].astype(str)
                        max_len = max(len(col) * 1.2, col_values.str.len().max() * 1.2) if len(col_values) > 0 else len(col) * 1.2
                        worksheet.set_column(idx, idx, max_len)
                    
                    # Write header with format
                    for col_num, value in enumerate(sheet_df.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                
                # Add charts
                chart_sheet = workbook.add_worksheet('Contribution Charts')
                
                # Create a chart for total contribution score
                chart1 = workbook.add_chart({'type': 'column'})
                chart1.add_series({
                    'name': 'Total Score',
                    'categories': ['Overall Summary', 1, df.columns.get_loc('username'), min(len(df), 10), df.columns.get_loc('username')],
                    'values': ['Overall Summary', 1, df.columns.get_loc('contribution_score'), min(len(df), 10), df.columns.get_loc('contribution_score')],
                })
                chart1.set_title({'name': 'Top 10 Contributors by Score'})
                chart1.set_x_axis({'name': 'Contributor'})
                chart1.set_y_axis({'name': 'Score'})
                chart_sheet.insert_chart('A1', chart1, {'x_scale': 2, 'y_scale': 1.5})
                
                # Organization comparison charts
                row_pos = 20
                for metric in ['pr_merged', 'commits', 'issues_opened']:
                    chart = workbook.add_chart({'type': 'column'})
                    
                    # Add a series for each organization
                    for i, org in enumerate(self.organizations):
                        prefix = f"{org.replace('-', '_')}_"
                        if f'{prefix}{metric}' in df.columns:
                            chart.add_series({
                                'name': f'{org}',
                                'categories': ['Overall Summary', 1, df.columns.get_loc('username'), min(len(df), 10), df.columns.get_loc('username')],
                                'values': ['Overall Summary', 1, df.columns.get_loc(f'{prefix}{metric}'), min(len(df), 10), df.columns.get_loc(f'{prefix}{metric}')],
                            })
                    
                    chart.set_title({'name': f'Top 10 Contributors by {metric.replace("_", " ").title()}'})
                    chart.set_x_axis({'name': 'Contributor'})
                    chart.set_y_axis({'name': 'Count'})
                    chart_sheet.insert_chart(f'A{row_pos}', chart, {'x_scale': 2, 'y_scale': 1.5})
                    row_pos += 20
            
            logger.info(f"Data exported successfully to {output_file}")
        except Exception as e:
            logger.error(f"Error exporting to Excel: {str(e)}")
            try:
                # Fallback to CSV export
                csv_file = output_file.replace('.xlsx', '.csv')
                df.to_csv(csv_file, index=False)
                logger.warning(f"Exported data to CSV instead: {csv_file}")
            except:
                logger.error("Failed to export data in any format.")
    
    def _paginated_request(self, url):
        """
        Make paginated requests to GitHub API.
        
        Args:
            url (str): GitHub API URL
            
        Returns:
            dict/list: Combined results from all pages
        """
        # Check cache first
        if url in self.cache:
            logger.debug(f"Using cached response for {url}")
            return self.cache[url]
            
        # Check rate limit before making request
        self.wait_for_rate_limit()
        
        results = None
        next_url = url
        page_count = 0
        max_retries = 3
        retry_count = 0
        
        while next_url:
            try:
                logger.debug(f"Making API request to: {next_url}")
                response = requests.get(next_url, headers=self.headers)
                
                # Update rate limit info from response headers
                if 'X-RateLimit-Remaining' in response.headers:
                    self.rate_limit_remaining = int(response.headers['X-RateLimit-Remaining'])
                if 'X-RateLimit-Reset' in response.headers:
                    self.rate_limit_reset = int(response.headers['X-RateLimit-Reset'])
                
                if response.status_code == 403 and 'rate limit exceeded' in response.text.lower():
                    wait_time = self.rate_limit_reset - time.time() + 5
                    if wait_time > 0:
                        logger.warning(f"Rate limit exceeded. Waiting for {wait_time:.1f} seconds...")
                        time.sleep(wait_time)
                        # Try the request again after waiting
                        return self._paginated_request(url)
                    else:
                        self.update_rate_limit_info()
                    
                if response.status_code != 200:
                    raise Exception(f"API request failed: {response.status_code} - {response.text}")
                
                retry_count = 0  # Reset retry count on successful response
                data = response.json()
                page_count += 1
                
                # Initialize results based on data type
                if results is None:
                    if isinstance(data, list):
                        results = []
                    elif isinstance(data, dict) and 'items' in data:
                        results = {'items': []}
                    else:
                        results = data
                        break  # If it's not a list or doesn't have items, we don't paginate
                
                # Combine results
                if isinstance(results, list) and isinstance(data, list):
                    results.extend(data)
                elif 'items' in results and 'items' in data:
                    results['items'].extend(data['items'])
                
                # Check for pagination links
                link_header = response.headers.get('Link', '')
                next_url = None
                
                for link in link_header.split(','):
                    if 'rel="next"' in link:
                        next_url = link[link.index('<') + 1:link.index('>')]
                        break
                
                # If we're continuing to next page, check rate limit
                if next_url:
                    logger.debug(f"Fetching page {page_count+1} for {url.split('?')[0]}...")
                    self.wait_for_rate_limit()
            except requests.exceptions.RequestException as e:
                logger.error(f"Network error: {str(e)}")
                if retry_count < max_retries:
                    retry_count += 1
                    wait_time = 2 ** retry_count  # Exponential backoff
                    logger.warning(f"Retrying in {wait_time} seconds... (Attempt {retry_count}/{max_retries})")
                    time.sleep(wait_time)
                    continue
                else:
                    logger.error(f"Max retries exceeded for {url}")
                    raise
        
        # Cache the result to avoid duplicating requests
        self.cache[url] = results
        return results


def main():
    # Load environment variables
    load_dotenv()
    
    print("GitHub GSoC Candidate Contribution Analyzer")
    print("Current Date and Time (UTC):", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    
    # Get config from .env file or prompt user
    github_token = os.getenv("GITHUB_TOKEN")
    if not github_token:
        github_token = input("Enter your GitHub Personal Access Token: ")
        if not github_token:
            logger.error("Error: GitHub token is required")
            sys.exit(1)
    
    # Get organizations from .env
    org_str = os.getenv("GITHUB_ORGANIZATIONS", "AOSSIE-Org,StabilityNexus")
    organizations = [org.strip() for org in org_str.split(",")]
    print(f"Analyzing contributions across the following organizations: {', '.join(organizations)}")
    
    # Get usernames from .env or file
    usernames = []
    env_usernames = os.getenv("GITHUB_USERNAMES")
    username_file = 'github_usernames.txt'
    
    if env_usernames:
        usernames = [name.strip() for name in env_usernames.split(",") if name.strip()]
        logger.info(f"Loaded {len(usernames)} usernames from environment variables")
    elif os.path.exists(username_file):
        with open(username_file, 'r') as f:
            usernames = [line.strip() for line in f if line.strip() and not line.startswith('#')]
        logger.info(f"Loaded {len(usernames)} usernames from {username_file}")
    else:
        # Prompt for usernames
        print("Enter GitHub usernames to analyze (one per line). Enter a blank line when done:")
        while True:
            username = input()
            if not username:
                break
            usernames.append(username)

    if not usernames:
        logger.error("No usernames provided. Exiting.")
        sys.exit(1)

    # Save usernames to file for future use
    with open(username_file, 'w') as f:
        f.write("# GitHub usernames for contribution analysis\n")
        for username in usernames:
            f.write(f"{username}\n")

    # Time period in months (default: all time - None)
    time_period = os.getenv("TIME_PERIOD")
    if not time_period:
        time_period = input("Enter time period to analyze (in months, press Enter for all time): ")
    
    time_period = int(time_period) if time_period and time_period.isdigit() else None
    
    # Output file
    output_file = os.getenv("OUTPUT_FILE", "github_contributions.xlsx")
    
    # Create analyzer and process data
    logger.info(f"Analyzing contributions for {len(usernames)} users across {len(organizations)} organizations...")
    analyzer = GitHubContributionAnalyzer(github_token, organizations, time_period)
    
    try:
        # First check rate limit to ensure we have enough requests available
        analyzer.update_rate_limit_info()
        estimated_requests = len(usernames) * len(organizations) * 5  # Rough estimate
        
        if analyzer.rate_limit_remaining < 100 or analyzer.rate_limit_remaining < estimated_requests * 0.1:
            logger.warning(f"Warning: Only {analyzer.rate_limit_remaining} API requests remaining!")
            logger.warning(f"Estimated requests needed: ~{estimated_requests}")
            proceed = input("Continue anyway? (y/n): ").lower() == 'y'
            if not proceed:
                logger.info("Exiting.")
                sys.exit(0)
                
        results = analyzer.analyze_users(usernames)
        
        # Export to Excel
        if results:
            analyzer.export_to_excel(results, output_file)
        else:
            logger.warning("No results to export. Please check GitHub API access and user contributions.")
    except KeyboardInterrupt:
        logger.info("\nProcess interrupted. Saving progress...")
        analyzer.save_progress()
        logger.info("You can resume the analysis later by running the script again.")
    except Exception as e:
        logger.error(f"An error occurred: {str(e)}", exc_info=True)
        analyzer.save_progress()
        logger.info("Progress saved. You can resume the analysis later.")


if __name__ == "__main__":
    main()
