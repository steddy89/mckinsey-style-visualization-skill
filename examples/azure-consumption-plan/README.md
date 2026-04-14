# Azure Consumption Plan - Example Scripts

Real-world example using McKinsey/BCG style visualization techniques to create
Azure cloud consumption plans for Universitas Terbuka (April 2026 - March 2027).

## Scripts

| File | Description |
|------|-------------|
| `gen_bcg_ppt.py` | BCG-style PowerPoint (9 slides) with green takeaway bars, "So what?" callouts, traffic lights |
| `gen_excel.py` | Formula-driven Excel workbook (7 sheets) with per-month projections and editable growth rate |

## Requirements

```
pip install python-pptx openpyxl
```

## Usage

```bash
# Generate BCG-style PowerPoint
python gen_bcg_ppt.py

# Generate Excel consumption plan
python gen_excel.py
```

## BCG Style Elements Used

- **Green brand color** (#00A651) throughout
- **Takeaway bars** at top of every slide with bold action statements
- **"So what?" callout boxes** with insight text
- **Traffic light indicators** (Green/Amber status)
- **Left-bar KPI cards** with thick colored accent bars
- **Phased implementation roadmap** (Phase 1/2/3)
- **Readiness assessment matrix**
- **Calibri font** (BCG standard)
- **CONFIDENTIAL footer** on every slide

## Data Source

Based on actual Azure ACR (Azure Consumed Revenue) data for Universitas Terbuka
FY26 (Jul 2025 - Apr 2026), with 20% growth projection and two new initiatives:

1. **AI Course Generation** - GPT-4o, AI Search, Speech, Cosmos DB
2. **Ujian Online Platform** - 4-region deployment (JKT, BDG, SMG, Lampung) with Reserved Instances
