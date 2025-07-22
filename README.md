# Packing Template Generator

A Python application for generating standardized packing lists and ITP (Inspection and Test Plan) documents. The application provides both a graphical user interface (GUI) and a command-line interface (CLI) for filling out template Excel files with user-provided data.

## Features

- **Dual Interface**: Choose between a user-friendly GUI or command-line operation
- **Multiple Templates**: Supports different template types including Packing Lists and ITP documents
- **Smart Merged Cell Handling**: Automatically handles Excel merged cells correctly
- **Input Validation**: Ensures all required fields are filled before generating documents
- **Batch Processing**: Process multiple documents from the command line

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/packing-template-generator.git
   cd packing-template-generator
   ```

2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### GUI Mode

Run the application with the graphical interface:

```bash
python main.py
```

### Command Line Mode

Generate documents directly from the command line:

```bash
python batch_fill.py TEMPLATE_TYPE EQUIPMENT CUSTOMER PURCHASE_ORDER DRAWING_NO SERIAL_NO [--date DATE] [--output OUTPUT_FILE]
```

Example:
```bash
python batch_fill.py "Packing List" "72.5kV Disc Motorised + Dual E/S Motorised" "ACME Corp" "PO12345" "DRW-001" "SN-2023-001" --date "21/07/2023" --output "output.xlsx"
```

## Template Structure

Templates should be placed in the project root directory. The application currently supports:

- Packing List templates:
  - 72.5kV Disc Motorised + Dual E/S Motorised
  - 33kV Disc Manual + Single E/S Manual
- ITP templates (placeholder implementation)

## Building the Application

To create a standalone executable:

```bash
pip install pyinstaller
pyinstaller --onefile --noconsole --icon=icon.ico main.py
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

For support, please open an issue in the GitHub repository.
