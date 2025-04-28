# GitHub GSoC Candidate Contribution Analyzer

A Python-based tool to analyze and compare GitHub contributions from potential Google Summer of Code candidates across multiple organizations.

## Overview

This tool helps GSoC mentors evaluate candidates by collecting comprehensive metrics on their GitHub activity within specified organizations. It generates an Excel report with detailed statistics and visualizations to aid in the selection process.

## Features

- **Multi-Organization Support**: Analyze contributions across multiple GitHub organizations (currently configured for AOSSIE-Org and StabilityNexus)
- **Comprehensive Metrics**: Track pull requests, commits, issues, code reviews, and more
- **Rate Limit Handling**: Smart handling of GitHub API rate limits with automatic waiting and resumption
- **Progress Tracking**: Save analysis progress to resume interrupted operations
- **Excel Reporting**: Generate detailed Excel reports with:
  - Overall contribution summaries
  - Organization-specific metrics
  - Comparative visualizations
  - Customizable scoring system

## Metrics Collected

- Pull requests (total, merged, open)
- Commits
- Issues (opened, closed)
- Code review activity
- Repository contribution counts
- Issue and PR comments
- Custom contribution score based on weighted metrics

## Requirements

- Python 3.6+
- Required Python packages:
  - requests
  - pandas
  - xlsxwriter
  - python-dateutil

## Installation

1. Clone this repository or download the script files
2. Install required packages:
   ```bash
   pip install requests pandas xlsxwriter python-dateutil
