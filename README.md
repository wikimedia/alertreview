# Alertreview

## Overview

Alertreview is a tool designed to automate the aggregation and analysis of alert information from various sources. By processing these alerts, Alertreview generates comprehensive alert review reports for specified time periods, aiding in the monitoring and management of alert data efficiently. This project is specifically tailored for use with Google Apps Script, leveraging JavaScript for its implementation.

## Features

- **Alert Aggregation:** Gathers alert data from multiple sources, including email alerts and incident reports from VictorOps.
- **Alert Analysis:** Processes and analyzes alert data to produce detailed reports, highlighting key metrics such as total alert counts and percentages of top alerts.
- **Custom Reporting:** Generates customized alert review reports for specified time periods, facilitating targeted analysis and review.
- **Caching Mechanism:** Utilizes caching to enhance performance and reduce redundant data processing.
- **Google Apps Script Integration:** Designed to run within the Google Apps Script environment, allowing for seamless integration with Google Workspace applications.

## Documentation

For detailed documentation on the project setup, usage guidelines, and additional resources, please visit our [Alert Review Wiki](https://wikitech.wikimedia.org/wiki/Alert_review).

## Prerequisites

- Google Apps Script environment
- Access to relevant alert data sources (e.g., Gmail for email alerts, VictorOps for incident reports)

## Installation

1. **Clone the Repository:**
    ```bash
    git clone https://gitlab.wikimedia.org/repos/sre/alertreview.git
    ```

2. **Install Dependencies:**
    Ensure you have Node.js installed on your system to manage project dependencies.
    ```bash
    npm install
    ```

3. **Google Apps Script Setup:**
    - Deploy the project within your Google Apps Script environment.
    - Set up the necessary script properties (`DOCUMENT_ID`, `VO_API_KEY`, `VO_API_ID`) in the Google Apps Script project.

## Usage

1. **Configuration:**
    - Update the global constants in `Code.js` to match your project's requirements (e.g., `DOCUMENT_ID`, `VO_API_KEY`, `VO_API_ID`).

2. **Running the Tool:**
    - Execute the `generateAlertAnalysisReport` function within the Google Apps Script editor to generate an alert review report.

## Development

- **Code Linting:** ESLint is configured to enforce the StandardJS style guide and ensure proper documentation via JSDoc.
- **Testing:** Run linting tests using:
    ```bash
    npm test
    ```

## Contributing

We welcome contributions to the Alertreview project! If you're interested in helping improve the project, please follow our [contribution guidelines](CONTRIBUTING.md).

## License

This project is licensed under the GPL-3.0 License - see the [LICENSE](LICENSE) file for details.
