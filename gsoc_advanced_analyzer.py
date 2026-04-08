import os
import asyncio
import aiohttp
import pandas as pd
import time
import logging
from datetime import datetime
from dateutil.relativedelta import relativedelta
import sys
import math
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
    def __init__(self, token, organizations, time_period_months=6):
        self.headers = {
            'Authorization': f'bearer {token}',
            'Content-Type': 'application/json'
        }
        self.organizations = organizations if isinstance(organizations, list) else [organizations]
        self.graphql_url = "https://api.github.com/graphql"
        
        # GSoC Window: usually last 6 months is most relevant
        self.since_date = (datetime.now() - relativedelta(months=time_period_months))
        since_date_str = self.since_date.strftime('%Y-%m-%d')
        self.time_filter = f"updated:>={since_date_str}"
        logger.info(f"Analysis Window: Since {since_date_str} (Updated)")

        self.analysis_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.detailed_contribution_logs = []

    async def _graphql_request(self, session, query, variables):
        try:
            async with session.post(self.graphql_url, json={'query': query, 'variables': variables}, headers=self.headers) as response:
                if response.status == 200:
                    data = await response.json()
                    if 'errors' in data:
                        # Filter out NOT_FOUND errors (e.g. user doesn't exist anymore)
                        errors = [e for e in data['errors'] if e.get('type') != 'NOT_FOUND']
                        if errors:
                            logger.error(f"GraphQL errors: {errors}")
                    return data
                elif response.status in (403, 429):
                    # Simple rate limit handling
                    reset_time = int(response.headers.get("x-ratelimit-reset", time.time() + 60))
                    wait_time = max(reset_time - time.time() + 5, 5)
                    logger.warning(f"Rate limit hit. Waiting {wait_time}s...")
                    await asyncio.sleep(wait_time)
                    return await self._graphql_request(session, query, variables)
                else:
                    text = await response.text()
                    logger.error(f"GraphQL Error {response.status}: {text}")
                    return None
        except Exception as e:
            logger.error(f"Request exception: {str(e)}")
            return None

    async def fetch_user_contributions(self, session, username, org):
        """Ultra-robust retrieval of all contribution types."""
        query = """
        query($searchQuery: String!, $cursor: String) {
            search(query: $searchQuery, type: ISSUE, first: 100, after: $cursor) {
                pageInfo { hasNextPage endCursor }
                nodes {
                    __typename
                    ... on PullRequest {
                        title state url createdAt mergedAt additions deletions changedFiles
                        repository { nameWithOwner }
                        commits { totalCount }
                        comments { totalCount }
                    }
                    ... on Issue {
                        title state url createdAt
                        repository { nameWithOwner }
                        comments { totalCount }
                    }
                }
            }
        }
        """
        
        async def run_search(q):
            nodes, cursor, has_next = [], None, True
            while has_next:
                data = await self._graphql_request(session, query, {"searchQuery": q, "cursor": cursor})
                if not data or 'data' not in data or not data['data'].get('search'):
                    break
                s_data = data['data']['search']
                nodes.extend([n for n in s_data['nodes'] if n])
                has_next = s_data['pageInfo']['hasNextPage']
                cursor = s_data['pageInfo']['endCursor']
            return nodes

        # Fetch PRs and Issues in parallel for efficiency
        tasks = [
            run_search(f"author:{username} org:{org} type:pr {self.time_filter}"),
            run_search(f"author:{username} org:{org} type:issue {self.time_filter}"),
            run_search(f"reviewed-by:{username} org:{org} type:pr {self.time_filter}"),
            run_search(f"commenter:{username} -author:{username} org:{org} {self.time_filter}")
        ]
        
        prs, issues, reviews, helped = await asyncio.gather(*tasks)
        
        logger.info(f"[{username} @ {org}] Found {len(prs)} PRs, {len(issues)} Issues, {len(reviews)} Reviews, {len(helped)} Participations")
        
        return {
            'prs': prs,
            'issues': issues,
            'reviews': reviews,
            'helped': helped
        }

    async def get_user_stats(self, session, username):
        stats = {
            'Username': username,
            'Total Score': 0,
            'Merged PRs': 0,
            'Open PRs': 0,
            'Issues Opened': 0,
            'Reviews Done': 0,
            'Participations': 0,
            'Consistency (Active Days)': 0,
            'Avg PR Complexity': 0,
            'Active Repos': set(),
            'PR Links': []
        }
        
        pr_complexities = []
        active_dates = set()

        for org in self.organizations:
            org_prefix = f"{org.replace('-', '_')}_"
            org_stats = {'Score': 0, 'Merged': 0}
            
            data = await self.fetch_user_contributions(session, username, org)

            # Process PRs
            for pr in data['prs']:
                state = pr['state'].upper()
                is_merged = pr.get('mergedAt') is not None or state == 'MERGED'
                repo = pr['repository']['nameWithOwner']
                stats['Active Repos'].add(repo)
                
                date = pr['createdAt'][:10]
                active_dates.add(date)
                
                # Complexity: log(lines) * log(files)
                complexity = math.log1p(pr.get('additions', 0) + pr.get('deletions', 0)) * math.log1p(pr.get('changedFiles', 0))
                pr_complexities.append(complexity)

                if is_merged:
                    stats['Merged PRs'] += 1
                    org_stats['Merged'] += 1
                    # Score: Base 25 + Complexity factor + Commits factor
                    score = 25 + (complexity * 3) + min(pr.get('commits', {}).get('totalCount', 0), 10)
                elif state == 'OPEN':
                    stats['Open PRs'] += 1
                    score = 10 + complexity
                else:
                    score = 2 # Closed but not merged (effort)
                
                org_stats['Score'] += score
                stats['PR Links'].append(f"{pr['url']} ({'MERGED' if is_merged else state})")
                
                self.detailed_contribution_logs.append({
                    'Username': username, 'Org': org, 'Repo': repo, 'Type': 'PR',
                    'State': 'MERGED' if is_merged else state, 'Title': pr['title'], 'URL': pr['url'],
                    'Score': round(score, 1), 'Date': date, 'Additions': pr.get('additions'), 'Deletions': pr.get('deletions')
                })

            # Process Issues
            for issue in data['issues']:
                stats['Issues Opened'] += 1
                org_stats['Score'] += 2
                active_dates.add(issue['createdAt'][:10])
                stats['Active Repos'].add(issue['repository']['nameWithOwner'])
                self.detailed_contribution_logs.append({
                    'Username': username, 'Org': org, 'Repo': issue['repository']['nameWithOwner'], 'Type': 'Issue',
                    'State': issue['state'], 'Title': issue['title'], 'URL': issue['url'],
                    'Score': 2, 'Date': issue['createdAt'][:10]
                })

            # Process Reviews
            for review in data['reviews']:
                stats['Reviews Done'] += 1
                org_stats['Score'] += 10
                # Note: Review nodes in search don't have createdAt directly available in the same way, 
                # but we can approximate or skip for consistency metric
                stats['Active Repos'].add(review['repository']['nameWithOwner'])
                self.detailed_contribution_logs.append({
                    'Username': username, 'Org': org, 'Repo': review['repository']['nameWithOwner'], 'Type': 'Review',
                    'State': 'DONE', 'Title': f"Reviewed: {review['title']}", 'URL': review['url'],
                    'Score': 10, 'Date': review['createdAt'][:10]
                })

            # Process Participations (Comments)
            for h in data['helped']:
                stats['Participations'] += 1
                org_stats['Score'] += 3
                stats['Active Repos'].add(h['repository']['nameWithOwner'])
                self.detailed_contribution_logs.append({
                    'Username': username, 'Org': org, 'Repo': h['repository']['nameWithOwner'], 'Type': 'Participation',
                    'State': 'COMMENTED', 'Title': f"Participated in: {h['title']}", 'URL': h['url'],
                    'Score': 3, 'Date': h['createdAt'][:10]
                })

            stats[f"{org_prefix}Score"] = round(org_stats['Score'], 1)
            stats[f"{org_prefix}Merged"] = org_stats['Merged']
            stats['Total Score'] += org_stats['Score']

        stats['Total Score'] = round(stats['Total Score'], 1)
        stats['Consistency (Active Days)'] = len(active_dates)
        stats['Avg PR Complexity'] = round(sum(pr_complexities)/len(pr_complexities) if pr_complexities else 0, 2)
        stats['Active Repos'] = ", ".join(sorted(list(stats['Active Repos'])))
        stats['PR Links'] = "\n".join(stats['PR Links'])
        
        return stats

    async def analyze_users_async(self, usernames):
        results = []
        async with aiohttp.ClientSession() as session:
            # Control concurrency
            sem = asyncio.Semaphore(5)
            
            async def bounded_get(u):
                async with sem:
                    try:
                        return await self.get_user_stats(session, u)
                    except Exception as e:
                        logger.error(f"Failed analysis for {u}: {str(e)}")
                        return None
            
            tasks = [bounded_get(u) for u in usernames]
            gathered = await asyncio.gather(*tasks)
            results = [g for g in gathered if g is not None]
        return results

    def export_to_excel(self, data, output_file="gsoc_candidates_ranking.xlsx"):
        if not data:
            logger.warning("No data to export.")
            return

        df_main = pd.DataFrame(data)
        df_main = df_main.sort_values('Total Score', ascending=False)
        
        df_logs = pd.DataFrame(self.detailed_contribution_logs)
        
        try:
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                df_main.to_excel(writer, sheet_name='Leaderboard', index=False)
                if not df_logs.empty:
                    df_logs.to_excel(writer, sheet_name='Detailed Logs', index=False)
                
                workbook = writer.book
                # Formatting
                header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
                for sheet in writer.sheets.values():
                    sheet.freeze_panes(1, 1)
                    # Simple column auto-fit
                    for i, col in enumerate(df_main.columns if sheet == writer.sheets['Leaderboard'] else df_logs.columns):
                        sheet.set_column(i, i, 20)
            
            logger.info(f"Report exported to {output_file}")
            print(f"SUCCESS: Report generated as {output_file}")
        except Exception as e:
            logger.error(f"Excel export failed: {str(e)}")
            df_main.to_csv("leaderboard_fallback.csv")

async def main_run():
    load_dotenv()
    token = os.getenv("GITHUB_TOKEN")
    orgs = [o.strip() for o in os.getenv("GITHUB_ORGS", "AOSSIE-Org,DjedAlliance,StabilityNexus").split(",") if o.strip()]
    
    if not token:
        print("ERROR: GITHUB_TOKEN not found in .env")
        return

    # Load usernames from cleaned file
    try:
        with open("ContributionsCalc/cleaned_usernames.txt", "r") as f:
            usernames = [line.strip() for line in f if line.strip()]
    except FileNotFoundError:
        print("ERROR: cleaned_usernames.txt not found.")
        return

    print(f"Analyzing {len(usernames)} candidates across {len(orgs)} organizations...")
    
    analyzer = AdvancedGSoCAnalyzer(token, orgs, time_period_months=6)
    results = await analyzer.analyze_users_async(usernames)
    
    if results:
        analyzer.export_to_excel(results)
    else:
        print("No results found.")

if __name__ == "__main__":
    asyncio.run(main_run())
