import os
import asyncio
import aiohttp
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
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('advanced_debug.log'),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger('gsoc_advanced_analyzer')


class AdvancedGSoCAnalyzer:
    def __init__(self, token, organizations, time_period_months=None):
        """
        Initialize the advanced analyzer with GitHub token and organizations.
        """
        self.headers = {
            'Authorization': f'bearer {token}',
            'Content-Type': 'application/json'
        }
        self.organizations = organizations if isinstance(
            organizations, list) else [organizations]
        self.graphql_url = "https://api.github.com/graphql"

        # Set time period filter if specified
        self.time_filter = ""
        if time_period_months:
            since_date = (
                datetime.now() - relativedelta(months=time_period_months)).strftime('%Y-%m-%d')
            self.time_filter = f" created:>={since_date}"
            logger.info(
                f"Setting time filter to {time_period_months} months (since {since_date})")
        else:
            logger.info("No time filter set. Fetching ALL-TIME contributions.")

        self.completed_users = set()
        self.progress_file = "advanced_analysis_progress.json"

        self.analysis_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    async def _graphql_request(self, session, query, variables):
        async with session.post(self.graphql_url, json={'query': query, 'variables': variables}, headers=self.headers) as response:
            if response.status == 200:
                data = await response.json()
                if 'errors' in data:
                    logger.error(f"GraphQL returned errors: {data['errors']}")
                return data
            elif response.status in (403, 429):
                reset_time = int(response.headers.get(
                    "x-ratelimit-reset", time.time() + 60))
                wait_time = max(reset_time - time.time() + 5, 5)
                logger.warning(
                    f"Rate limit exceeded. Waiting {wait_time:.1f} seconds...")
                await asyncio.sleep(wait_time)
                return await self._graphql_request(session, query, variables)
            else:
                text = await response.text()
                logger.error(f"GraphQL Error {response.status}: {text}")
                return None

    async def fetch_user_issues_prs(self, session, username, org):
        """Fetch all PRs and issues for a user in an org using GraphQL."""
        query = """
        query($searchQuery: String!, $cursor: String) {
            search(query: $searchQuery, type: ISSUE, first: 100, after: $cursor) {
                pageInfo {
                    hasNextPage
                    endCursor
                }
                nodes {
                    ... on PullRequest {
                        __typename
                        title
                        state
                        createdAt
                        mergedAt
                        additions
                        deletions
                        repository {
                            nameWithOwner
                        }
                    }
                    ... on Issue {
                        __typename
                        title
                        state
                        createdAt
                        closedAt
                        repository {
                            nameWithOwner
                        }
                    }
                }
            }
        }
        """
        # Note: If time_filter is empty, this fetches all history.
        results = {'prs': [], 'issues': []}

        for type_filter in ["is:issue", "is:pr"]:
            search_query = f"author:{username} org:{org} {type_filter}{self.time_filter}"
            cursor = None
            has_next = True

            while has_next:
                variables = {"searchQuery": search_query, "cursor": cursor}
                data = await self._graphql_request(session, query, variables)
                if not data or 'data' not in data or not data['data'].get('search') or not data['data']['search'].get('nodes'):
                    break

                search_data = data['data']['search']
                for node in search_data.get('nodes', []):
                    if not node:
                        continue
                    if node['__typename'] == 'PullRequest':
                        results['prs'].append(node)
                    elif node['__typename'] == 'Issue':
                        results['issues'].append(node)

                page_info = search_data['pageInfo']
                has_next = page_info['hasNextPage']
                cursor = page_info['endCursor']

        return results

    async def get_user_stats(self, session, username):
        """Analyze a single user asynchronously."""
        logger.info(f"Analyzing {username}...")
        stats = {
            'username': username,
            'pr_total': 0, 'pr_merged': 0, 'pr_open': 0, 'pr_closed': 0,
            'issues_opened': 0, 'issues_closed': 0,
            'complexity_score': 0,
            'pr_details': [],
            'repos_contributed': set(),
            'organizations': {}
        }

        for org in self.organizations:
            org_stats = {
                'pr_total': 0, 'pr_merged': 0, 'pr_open': 0, 'pr_closed': 0,
                'issues_opened': 0, 'issues_closed': 0,
                'complexity_score': 0,
                'repos_contributed': set()
            }
            stats['organizations'][org] = org_stats

            data = await self.fetch_user_issues_prs(session, username, org)

            for pr in data['prs']:
                stats['pr_total'] += 1
                org_stats['pr_total'] += 1

                repo_name = pr['repository']['nameWithOwner'] if pr.get(
                    'repository') else 'unknown'
                stats['repos_contributed'].add(repo_name)
                org_stats['repos_contributed'].add(repo_name)

                state = pr['state']
                if state == 'OPEN':
                    stats['pr_open'] += 1
                    org_stats['pr_open'] += 1
                elif state == 'MERGED':
                    stats['pr_merged'] += 1
                    org_stats['pr_merged'] += 1
                elif state == 'CLOSED':
                    stats['pr_closed'] += 1
                    org_stats['pr_closed'] += 1

                additions = pr.get('additions', 0)
                deletions = pr.get('deletions', 0)

                # Complexity score: Base 1 point per PR + 1 point per 10 lines changed (capped at 100 points via lines)
                pr_complexity = 1.0 + min(additions + deletions, 1000) / 10.0
                stats['complexity_score'] += pr_complexity
                org_stats['complexity_score'] += pr_complexity

                stats['pr_details'].append(
                    f"[{repo_name}] {pr['title']} (State: {state}, +{additions}/-{deletions})")

            for issue in data['issues']:
                stats['issues_opened'] += 1
                org_stats['issues_opened'] += 1

                repo_name = issue['repository']['nameWithOwner'] if issue.get(
                    'repository') else 'unknown'
                stats['repos_contributed'].add(repo_name)
                org_stats['repos_contributed'].add(repo_name)

                if issue['state'] == 'CLOSED':
                    stats['issues_closed'] += 1
                    org_stats['issues_closed'] += 1

            org_stats['repos_contributed'] = len(
                org_stats['repos_contributed'])

        stats['repos_contributed'] = len(stats['repos_contributed'])
        # Store as string for Excel
        stats['pr_details'] = "\n".join(
            stats['pr_details']) if stats['pr_details'] else "None"

        # Flatten for easy excel export
        for org in self.organizations:
            prefix = f"{org.replace('-', '_')}_"
            for k, v in stats['organizations'][org].items():
                stats[f"{prefix}{k}"] = v

        del stats['organizations']

        logger.info(
            f"Completed analysis for {username}: Score {stats['complexity_score']:.1f}, {stats['pr_merged']} Merged PRs")
        return stats

    async def analyze_users_async(self, usernames):
        results = []
        async with aiohttp.ClientSession() as session:
            sem = asyncio.Semaphore(5)

            async def bounded_get(u):
                async with sem:
                    try:
                        return await self.get_user_stats(session, u)
                    except Exception as e:
                        logger.error(f"Failed analysis for {u}: {e}")
                        return None

            tasks = [bounded_get(u) for u in usernames]
            gathered = await asyncio.gather(*tasks)
            results = [g for g in gathered if g is not None]

        return results

    def analyze_users(self, usernames):
        return asyncio.run(self.analyze_users_async(usernames))

    def export_to_excel(self, data, output_file="gsoc_advanced_report.xlsx"):
        if not data:
            logger.warning("No data to export.")
            return

        df = pd.DataFrame(data)

        # Advanced weighting formula
        df['total_gsoc_score'] = (
            (df['pr_merged'] * 10) +
            (df['pr_open'] * 3) +
            (df['issues_opened'] * 2) +
            (df['issues_closed'] * 4) +
            df['complexity_score']
        )

        df = df.sort_values('total_gsoc_score', ascending=False)

        metadata = {
            'Generated On': self.analysis_date,
            'Organizations Analyzed': ', '.join(self.organizations),
            'Time Period': "All time" if not self.time_filter else f"Last {self.time_filter}",
            'Number of Contributors': len(df)
        }

        try:
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                pd.DataFrame(list(metadata.items()), columns=['Metric', 'Value']).to_excel(
                    writer, sheet_name='Info', index=False)
                df.to_excel(writer, sheet_name='GSoC Insights', index=False)

            logger.info(
                f"Data exported successfully to {output_file} with Advanced Metrics!")
        except Exception as e:
            logger.error(f"Error exporting to Excel: {e}")


if __name__ == "__main__":
    load_dotenv()
    token = os.getenv("GITHUB_TOKEN")
    orgs = [o.strip() for o in os.getenv(
        "GITHUB_ORGS", "").split(',') if o.strip()]
    users = [u.strip() for u in os.getenv(
        "GITHUB_USERS", "").split(',') if u.strip()]

    if not token or not orgs:
        logger.error("Please set GITHUB_TOKEN and GITHUB_ORGS in .env")
        sys.exit(1)

    # Deliberately setting time_period_months=None to fetch ALL-TIME history
    # to avoid the "0 score" issue if users haven't contributed in the last 6 months.
    analyzer = AdvancedGSoCAnalyzer(token, orgs, time_period_months=None)
    results = analyzer.analyze_users(users)
    analyzer.export_to_excel(results)
