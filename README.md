![Python](https://img.shields.io/badge/Python-3.10-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Status](https://img.shields.io/badge/Status-Production_Ready-brightgreen)# Risk Analysis Platform

## Monte Carlo Risk Analysis for Mega Infrastructure Projects

### Bachelor Thesis | Underground Metro Station Case Study

---

# Overview

Risk Analysis Platform is a professional probabilistic risk analysis tool developed as the primary original contribution of a bachelor thesis focused on mega infrastructure projects in the Middle East construction sector.

The platform integrates:

- Monte Carlo Simulation
- Correlated construction risk modeling
- AI-based risk prediction
- Primavera P6 schedule integration
- Real-time weather data
- Professional reporting and visualization

The system was designed specifically for underground metro station construction projects and large-scale infrastructure developments.

---

# Key Features

## Monte Carlo Simulation Engine
- 10,000 simulation iterations
- Latin Hypercube Sampling (LHS)
- Correlated probabilistic risk modeling
- Statistical confidence analysis (P50 / P80 / P90)

## Correlated Risk Factors
The platform models 13 major construction risks:

1. Weather Risk  
2. Labor Productivity Risk  
3. Material Supply Risk  
4. Ground Condition Risk  
5. Equipment Failure Risk  
6. War & Geopolitical Risk  
7. Pandemic Risk  
8. Supply Chain Disruption  
9. Regulatory Delay Risk  
10. Financial Risk  
11. Design Change Risk  
12. Safety Incident Risk  
13. Extreme Event Risk  

## Risk Control Logic
- `MAX_DURATION_FACTOR = 1.10`
- Prevents unrealistic compounded delays
- Improves realism of simulation outputs

## AI Risk Prediction
- Random Forest Machine Learning model
- Trained using 10,000 synthetic project scenarios
- Predicts schedule and cost risk exposure

## Primavera P6 Integration
- Parse Primavera XER files
- Import Excel schedule exports
- Support deterministic baseline schedules

## Real-Time Weather Integration
- Open-Meteo API integration
- Current environmental conditions
- Weather-based schedule influence

## Professional Reporting
- Automated Excel report generation
- Styled worksheets and dashboards
- Statistical summaries and charts

## Interactive Dashboard
- Streamlit-based web application
- Interactive risk visualization
- User-friendly interface
![Risk Analysis Platform Dashboard](screenshot.png)
(Then upload a screenshot of your software)



---

# Simulation Results

| Metric | Value |
|---|---|
| P50 Duration | 1,892 Days |
| P80 Duration | 1,971 Days |
| P90 Duration | 2,015 Days |
| P50 Cost | \$68,825,911 |
| P80 Cost | \$71,072,135 |

---

# Technologies Used

| Technology | Purpose |
|---|---|
| Python 3.10 | Core development |
| Streamlit | Web dashboard |
| NumPy | Numerical computation |
| Pandas | Data processing |
| Plotly | Interactive visualization |
| SciPy | Statistical analysis |
| scikit-learn | AI prediction |
| OpenPyXL | Excel report generation |
| Requests | API integration |

---

# Installation

## Clone the Repository

```bash
git clone https://github.com/abdullahtarek20/Risk-Analysis-Platform.git
cd Risk-Analysis-Platform
```

## Install Dependencies

```bash
pip install streamlit pandas numpy plotly scikit-learn scipy openpyxl requests
```

---

# How to Run

```bash
streamlit run risk_analysis_platform.py
```

After running the command, Streamlit will automatically open the web dashboard in your browser.

---

# Repository Structure

```text
Risk-Analysis-Platform/
│
├── risk_analysis_platform.py     # Main application
├── README.md                     # Project documentation
│
└── outputs/                      # Generated reports and exports
```

---

# Thesis Reference

This software was developed as part of the following bachelor thesis:

> **"Probabilistic Risk Analysis for Mega Infrastructure Projects: A Monte Carlo Simulation Approach with Correlated Risk Factors"**

### Case Study
Underground Metro Station Construction in the Middle East

---

# Academic Information

| Item | Details |
|---|---|
| Author | Abdullah Tarek |
| Degree | Bachelor of Science in Civil Engineering |
| Field | Civil Engineering |
| Year | 2026 |

---

# Research Contribution

This project contributes to the field of construction risk management by combining:

- Probabilistic schedule analysis
- Correlated infrastructure risk modeling
- AI-assisted prediction methods
- Practical Primavera P6 integration
- Real-world Middle East construction risks

The platform aims to support decision-making for large-scale infrastructure projects with improved schedule and cost forecasting accuracy.

---

# License

MIT License

Copyright (c) 2026 Abdullah Tarek

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

---

# Author

**Abdullah Tarek**  
Bachelor Thesis in Civil Engineering  
2026

---

# Contact

For academic or research-related inquiries regarding this thesis or software, please contact the author through the university.

---

# Last Updated

May 2026
