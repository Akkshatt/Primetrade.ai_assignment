# Cryptocurrency Live Data Fetcher

## Overview

This project fetches live cryptocurrency data for the top 50 cryptocurrencies by market capitalization using the CoinGecko API. The data includes the cryptocurrency name, symbol, current price in USD, market capitalization, 24-hour trading volume, and 24-hour price change percentage. The data is then analyzed and presented in a live-updating Excel sheet that updates every 5 minutes.

## Features

- Fetches live cryptocurrency data using the CoinGecko API.
- Performs basic data analysis including:
  - Identifying the top 5 cryptocurrencies by market cap.
  - Calculating the average price of the top 50 cryptocurrencies.
  - Analyzing the highest and lowest 24-hour percentage price change among the top 50.
- Updates an Excel sheet with the latest data every 5 minutes.

## Requirements

- Python 3.x
- `requests` library
- `pandas` library
- `openpyxl` library

## Installation

1. **Clone the repository:**

   ```bash
   git clone https://github.com/yourusername/crypto-live-data-fetcher.git
   cd crypto-live-data-fetcher
