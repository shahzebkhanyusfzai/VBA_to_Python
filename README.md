# Insurance Policy Fee Projection Model

## Overview
This project transforms an existing VBA-based model for projecting monthly fees from a block of insurance policies into a more efficient Python application, utilizing xlwings to maintain integration with Excel. The original goal was to enhance the model's performance, reducing processing time from minutes to seconds, while still leveraging the familiar Excel interface for input and output manipulation.

## Goal
The primary objective of this project is to project the monthly fees to be earned from a block of insurance policies over a span of 600 months (50 years) across various stochastic scenarios. The model calculates the total projected fees by month and scenario, outputting the results in a CSV file for further analysis.

## Methodology
The Python model, like its VBA predecessor, loops through numerous stochastic scenarios to project forward the policies and determine their status (active or lapsed) at each point. The methodology involves:
- Looping through each month to assess policy survival based on lapse rates and mortality, adjusted for policyholder age and policy duration.
- Incorporating pandemic effects by increasing mortality rates by a specified severity factor for the pandemic year and by 50% of that factor for the following year.
- Drawing uniform random numbers to decide policy lapse or survival, influenced by the lapse and mortality assumptions.

## Inputs
- **Inforce Data:** A list of active policies, including specific lapse tables, mortality tables, and monthly policy fees.
- **Assumptions:** Details on lapse and mortality assumptions, pandemic incidence, and severity factors.

## Output
A CSV file detailing the scenarios and total fees for each of the 600 months, resulting in a comprehensive (scen X 600) output table.

## Quick Start Guide

### Prerequisites
Ensure you have Python and xlwings installed. xlwings allows for seamless integration with Microsoft Excel, enabling Python to interact with Excel workbooks.

### Running the Script
1. Open the Excel workbook containing the 'Inforce' and 'Assumptions' tabs with the necessary input data.
2. Execute the Python script. The script reads the input data from the Excel workbook, performs the projections based on the described methodology, and writes the output to a CSV file.
3. The output CSV file will be saved in the specified directory, containing the total projected fees by month and scenario.

## Benefits of Transitioning from VBA to Python

This project showcases a significant enhancement in processing efficiency by migrating from Excel's VBA to Python, leveraging xlwings for Excel integration. Key benefits include:

- **Increased Speed:** Processing times were dramatically reduced from 7-10 minutes to under 30 seconds, enhancing productivity and enabling faster decision-making.
- **Scalability:** Python's scalability allows for more complex analyses and larger datasets without a significant increase in processing time.
- **Advanced Libraries:** Utilizing Python's extensive libraries, such as Pandas for data manipulation and NumPy for numerical analysis, contributes to the model's efficiency and capabilities.
- **Cross-Platform Compatibility:** Python's ability to run on various operating systems provides flexibility in deployment and use across different platforms.
- **Community Support:** The vast Python community offers a wealth of resources, contributing to ongoing improvements and innovation in project development.

Choosing Python over VBA for this project not only achieved the immediate goal of faster execution times but also opened the door to future enhancements and more complex modeling capabilities.

