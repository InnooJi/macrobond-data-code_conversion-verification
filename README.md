# macrobond-data-code_conversion-verification
# Macrobond Data Code Verification

## Overview
The **Macrobond Data Code Verification** project provides a comprehensive solution for automating the validation of foreign code mappings to Macrobond data series. This tool ensures accuracy, efficiency, and reliability in transitioning clients from legacy systems to Macrobond's robust data infrastructure. By leveraging Python, Macrobond Data API, and Excel integration, this project addresses the challenges of manual mapping and data integrity verification.

---

## Problem Statement
Many Macrobond clients transition from legacy data providers, such as Haver, and require accurate mapping of their existing data codes to Macrobond's expressions. This process typically serves two primary purposes:

1. **Proof of Concept**: Demonstrating to potential clients that their critical data series can be accurately mapped to Macrobond's systems.
2. **Seamless Onboarding**: Ensuring smooth transitions for new clients by automating bulk data series conversions.

The current manual workflow, though functional, is prone to:
- Human errors, including incorrect mappings or missed series.
- Delays due to the high volume of requests and the manual validation process.

Efficient and accurate validation of these mappings is essential to establish trust and maintain client relationships.

---

## High-Level Solution
This project introduces an automated Python-based tool designed to:
- Fetch the most recent data series from Macrobond.
- Compare the data against client-provided Excel sheets to validate mappings.
- Identify discrepancies in data series, including mismatches in dates or values.
- Resample data to match different frequencies, such as weekly, ensuring alignment with client expectations.

Key features include:
- **Data Alignment**: Handles weekly, daily, and other frequencies with precision.
- **Enhanced Debugging**: Provides detailed logging for mismatches and potential errors.
- **Unit Conversion Handling**: Resolves discrepancies caused by scale differences in values.
- **Automated Results Reporting**: Generates an Excel sheet summarizing validation results.

---

## Implementation Details
1. **Weekly Data Fetching**: Efficiently retrieves and resamples data to align with client-required frequencies (e.g., weekly).
2. **Value Comparison**: Robust logic for comparing expected and actual values with support for scale adjustments and missing data.
3. **Validation Workflow**: Automates the end-to-end validation process, ensuring accurate foreign code mappings.
4. **Result Reporting**: Clear and structured result logs, detailing mismatches and successes.

---

## Workflow
1. **Input**:
   - An Excel workbook containing:
     - A frequency-specific sheet (e.g., `WEEKLY`) with foreign codes and Macrobond codes.
     - A data sheet (`ALL TEAMS - W`) containing actual series data for comparison.

2. **Processing**:
   - Fetches weekly series from the Macrobond API.
   - Resamples and aligns data.
   - Validates against the Excel sheet.

3. **Output**:
   - An updated Excel sheet with:
     - Validation status (Success, Mismatch, Not Found).
     - Accuracy percentage for each foreign code.

