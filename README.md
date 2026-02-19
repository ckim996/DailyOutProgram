# ShipStation Logistics & Picking Optimizer

An automated fulfillment engine designed to streamline warehouse operations and minimize shipping costs. This tool integrates with the ShipStation API to automate carrier selection and generate optimized picking lists.

## Key Technical Features
* **Concurrent Rate Fetching**: Utilizes `ThreadPoolExecutor` to handle parallel API requests, significantly reducing the time required to fetch and compare live shipping rates.
* **Picking List Algorithm**: A custom-built logic that organizes parts into a specific 40/41-row alternating grid, optimized for physical warehouse document layouts.
* **Automated Shipping Decisions**: Evaluates orders to determine the best carrier (UPS vs. USPS) and packaging based on weight, dimensions, and destination state.
* **Financial Tracking**: Automatically calculates potential savings by comparing real-time shipping costs against internal database benchmarks.
* **Inventory Integration**: Cross-references orders with local inventory data to identify out-of-stock items and flag potential "eBay purchases".

## Tech Stack
* **Language**: Python
* **Libraries**: `pandas`, `openpyxl` (Excel automation), `concurrent.futures` (Concurrency)
* **API**: ShipStation V1

## Project Structure
* `main.py`: Main execution logic and Excel workbook generation.
* `src/shipping/`: Logic for the shipping engine and rate optimizer.
* `src/lookup/`: SKU and Part number lookup utilities.