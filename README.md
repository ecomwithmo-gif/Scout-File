# Excel Formatter Pro 2.0

A modern, feature-rich Excel processing application with advanced analytics, streaming support, and a beautiful user interface.

## 🚀 Features

### Core Functionality
- **Advanced Excel Processing**: Process large Excel files with intelligent streaming
- **Data Validation**: Comprehensive validation before processing to prevent errors
- **Memory Optimization**: Efficient memory usage with chunked processing and caching
- **Progress Tracking**: Real-time progress updates with cancellation support

### Analytics & Insights
- **Interactive Charts**: Profit distribution, sales rank analysis, and more
- **Data Statistics**: Comprehensive statistical analysis of your data
- **Smart Insights**: AI-powered insights and recommendations
- **Export Capabilities**: Export analytics to Excel format

### Modern UI/UX
- **Clean Interface**: Modern, responsive design with dark/light themes
- **Drag & Drop**: Easy file upload with drag-and-drop support
- **Keyboard Shortcuts**: Power user features with keyboard shortcuts
- **Real-time Feedback**: Live status updates and progress tracking

### Performance Features
- **Streaming Processing**: Handle files of any size with streaming
- **Calculation Caching**: Intelligent caching for faster repeated calculations
- **Memory Management**: Automatic memory cleanup and optimization
- **Background Processing**: Non-blocking UI with background processing

## 📦 Installation

### Prerequisites
- Python 3.8 or higher
- Windows, macOS, or Linux

### Quick Start
1. Clone the repository:
```bash
git clone https://github.com/your-username/excel-formatter-pro.git
cd excel-formatter-pro
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
python main.py
```

### Development Setup
```bash
# Install development dependencies
pip install -e ".[dev]"

# Run with development tools
python -m pytest
black .
isort .
flake8 .
```

## 🎯 Usage

### Basic Usage
1. **Launch the application** by running `python main.py`
2. **Upload your main Excel file** using the file upload section
3. **Optionally upload a cost/MSRP file** for enhanced calculations
4. **Configure settings** like shipping costs and processing options
5. **Click "Process Excel File"** to start processing
6. **View analytics** in the analytics panel
7. **Export results** or analytics as needed

### Advanced Features

#### Streaming Processing
For large files (>50MB), the application automatically uses streaming processing:
- Files are processed in chunks for better memory management
- Progress is tracked in real-time
- Processing can be cancelled at any time

#### Data Validation
Before processing, the application validates:
- File format and accessibility
- Required columns presence
- Data type consistency
- Data integrity checks

#### Analytics Dashboard
The analytics panel provides:
- **Overview**: Basic dataset information and data quality metrics
- **Charts**: Interactive visualizations of your data
- **Statistics**: Detailed statistical analysis
- **Insights**: AI-powered recommendations and insights

### Keyboard Shortcuts
- `Ctrl+O`: Open main file
- `Ctrl+Shift+O`: Open secondary file
- `Ctrl+R`: Start processing
- `Ctrl+Q`: Quit application
- `Ctrl+Shift+T`: Toggle theme
- `F11`: Toggle fullscreen
- `Escape`: Exit fullscreen

## ⚙️ Configuration

### Settings File
Configuration is stored in `config/settings.json`:

```json
{
  "app": {
    "name": "Excel Formatter Pro",
    "version": "2.0.0",
    "theme": {
      "default": "light",
      "colors": {
        "primary": "#3b82f6",
        "success": "#10b981",
        "error": "#ef4444"
      }
    }
  },
  "processing": {
    "chunk_size": 1000,
    "max_file_size_mb": 500,
    "enable_streaming": true,
    "enable_caching": true
  }
}
```

### Header Mappings
Column mappings are defined in `config/header_mappings.json`:

```json
{
  "header_mappings": {
    "ASIN": "ASIN",
    "Title": "Title",
    "UPC": "UPC",
    "Sales Rank: Current": "Sales Rank"
  },
  "conditional_formatting": {
    "sales_rank_ranges": {
      "excellent": {"min": 0, "max": 150000, "color": "90EE90"},
      "good": {"min": 150001, "max": 500000, "color": "FFD580"},
      "poor": {"min": 500001, "max": 999999999, "color": "FFB6B6"}
    }
  }
}
```

## 🏗️ Architecture

### Project Structure
```
excel-formatter-pro/
├── main.py                 # Main application entry point
├── requirements.txt        # Python dependencies
├── setup.py               # Package setup
├── config/                # Configuration files
│   ├── settings.json
│   └── header_mappings.json
├── core/                  # Core processing logic
│   └── excel_processor.py
├── ui/                    # User interface
│   ├── main_window.py
│   └── components/
├── analytics/             # Analytics and visualization
│   └── data_analytics.py
└── utils/                 # Utilities and helpers
    ├── config_manager.py
    ├── data_validator.py
    └── exceptions.py
```

### Key Components

#### ExcelProcessor
- Handles all Excel file processing
- Implements streaming for large files
- Manages memory optimization
- Provides progress callbacks

#### DataValidator
- Validates files before processing
- Checks data integrity
- Provides detailed validation reports
- Prevents processing errors

#### DataAnalytics
- Generates comprehensive statistics
- Creates interactive charts
- Provides data insights
- Exports analytics data

#### MainWindow
- Modern UI with responsive design
- Manages all user interactions
- Coordinates between components
- Handles theme and settings

## 🔧 Development

### Adding New Features

1. **Create feature branch**:
```bash
git checkout -b feature/new-feature
```

2. **Implement feature** following the existing architecture patterns

3. **Add tests** for new functionality

4. **Update documentation** as needed

5. **Submit pull request**

### Code Style
- Follow PEP 8 guidelines
- Use type hints for all functions
- Add docstrings for all classes and methods
- Use meaningful variable and function names

### Testing
```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=.

# Run specific test file
pytest tests/test_excel_processor.py
```

## 🐛 Troubleshooting

### Common Issues

#### Memory Issues
- Reduce chunk size in settings
- Close other applications
- Use streaming mode for large files

#### File Processing Errors
- Check file format (must be .xlsx)
- Ensure file is not open in another application
- Verify file permissions

#### UI Issues
- Try switching themes
- Restart the application
- Check console for error messages

### Logs
Application logs are stored in `excel_formatter.log` for debugging.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests
5. Submit a pull request

## 📞 Support

For support, please:
1. Check the troubleshooting section
2. Search existing issues
3. Create a new issue with detailed information

## 🎉 Acknowledgments

- Built with [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) for the modern UI
- Uses [Pandas](https://pandas.pydata.org/) for data processing
- Powered by [Matplotlib](https://matplotlib.org/) and [Seaborn](https://seaborn.pydata.org/) for analytics
- Excel processing with [OpenPyXL](https://openpyxl.readthedocs.io/)

---

**Excel Formatter Pro 2.0** - Transform your Excel data processing workflow! 🚀