# ILS Kickoff Analysis Tool

Generate banking OD/NSF analysis reports from client data. Produces formatted Excel workbooks and PowerPoint presentations with charts and tables.

## Quick Start

1. **Place your CSV file** in the `data/` folder
2. **Edit `config.yaml`** -- set `data_file` to your CSV filename:
   ```yaml
   data_file: "data/1774_INB_OD.Tran.Combo_20260203 (1).csv"
   ```
3. **Run the analysis:**
   ```
   make run
   ```

Reports are saved to the `output/` folder.

## Setup

### macOS/Linux

```bash
make setup
```

### Windows

Double-click `setup.bat`, then double-click `run.bat` to run.

### Manual Setup

```bash
python -m venv .venv
source .venv/bin/activate    # macOS/Linux
.venv\Scripts\activate       # Windows
pip install -r requirements.txt
```

## Usage

### With config file (recommended)

```bash
cp config.example.yaml config.yaml
# Edit config.yaml to set your data_file path
python -m ils_kickoff
```

### With command-line arguments

```bash
python -m ils_kickoff data/your_file.csv
python -m ils_kickoff data/your_file.csv --output ./reports/ --client-id 1774
python -m ils_kickoff data/your_file.csv --verbose
```

### Options

```
python -m ils_kickoff --help
```

| Option | Description |
|--------|-------------|
| `data_file` | Path to client CSV/Excel file |
| `--config`, `-c` | Path to config.yaml (default: config.yaml) |
| `--output`, `-o` | Output directory (default: output/) |
| `--client-id` | Client identifier (auto-extracted from filename) |
| `--client-name` | Client name for report titles |
| `--verbose`, `-v` | Enable debug logging |

## Configuration

Copy `config.example.yaml` to `config.yaml` and edit:

- **`data_file`** (required): Path to your CSV file
- **`client_id`** (optional): Auto-extracted from filename if omitted
- **`output_dir`** (optional): Where to save reports (default: `output/`)
- **`outputs`**: Toggle Excel, PowerPoint, or HTML chart generation

See `config.example.yaml` for all available options.

## Output

The tool generates:

- **Excel workbook** with a cover sheet, table of contents, and 15 formatted analysis tabs with charts
- **PowerPoint deck** with chart images and data tables on each slide

### Analyses

1. Account Status Summary (All Accounts)
2. Account Type (Open Accounts)
3. Personal Deposit Distribution
4. Business Deposit Distribution
5. Personal NSF Stratification (Volume)
6. Business NSF Stratification (Volume)
7. Personal NSF Pay Ratio
8. Business NSF Pay Ratio
9. Personal NSF Full Behavioral Metrics
10. Business NSF Full Behavioral Metrics
11. Personal OD Status Stratification
12. Business OD Status Stratification
13. Reg E Distribution
14. OD Limit Stratification
15. Historical Reg E by Year Opened

## Requirements

- Python 3.10+
- Dependencies listed in `requirements.txt`

## Development

```bash
make setup-dev
make test
make lint
```
