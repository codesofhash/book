# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build Commands

```bash
# Build the project
dotnet build

# Run the application
dotnet run

# Clean and rebuild
dotnet clean && dotnet build
```

## Project Overview

CSharpFlexGrid is a Windows Forms desktop application (.NET 9.0) for processing advertising/media booking orders from Excel files and managing them through an interactive calendar grid interface. It handles booking data from multiple formats (Al Mulla and Direct NCCAL Excel formats).

## Architecture

### Application Flow

```
Program.cs (entry point)
    └→ SplashScreen (3-second splash)
        └→ MainMenuForm (navigation hub)
            ├→ BookingManagerForm (file selection & Excel processing)
            └→ CalendarGridForm (calendar-based booking view)
```

### Core Components

- **BookingOrderReader.cs**: Parses Excel files using ExcelDataReader, auto-detects sheet format, extracts booking data into structured models
- **BookingOrderModels.cs**: Data contracts - `BookingOrder` (root entity with agency, advertiser, campaign period, spots), `Spot` (individual ad with dates/times), `CampaignPeriod` (date range)
- **DatabaseHelper.cs**: MySQL operations using MySqlConnector, connection via environment variables
- **CustomDataGridView.cs**: Extended DataGridView with copy/cut/paste, row/column insertion and deletion
- **EditableAutoCompleteComboBox.cs**: Autocomplete 
- box that overlays DataGridView cells

### Configuration Files

- `.env`: Database credentials (`DB_HOST`, `DB_USER`, `DB_PASSWORD`, `DB_NAME`, `DB_PORT`) and `JSON_OUTPUT_PATH` for processed orders
- `settings.json`: Application settings (`DefaultGridMode`, `DealSearchDays`)

### Key Dependencies

- ExcelDataReader / ExcelDataReader.DataSet - Excel file parsing
- MySqlConnector - Database connectivity
- Newtonsoft.Json - JSON serialization
- DotNetEnv - Environment variable management

## Code Patterns

- Single namespace: `CSharpFlexGrid`
- Forms follow pattern: `{Name}Form.cs` with `{Name}Form.Designer.cs`
- Database credentials loaded from `.env` in `DatabaseHelper` constructor
- Processed booking orders saved as JSON to `ProcessedOrders/` directory with timestamp naming
