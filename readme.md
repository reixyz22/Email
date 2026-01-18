# Email Automation System

Automated email outreach system built for Chi-AWE (Chicago Asian Women Empowerment), a 501(c)3 nonprofit organization. Reduces manual email labor by **30+ hours annually** through CSV-driven personalized messaging.

## Overview

This Python automation tool streamlines nonprofit email campaigns by:
- Reading recipient data from CSV files
- Converting DOCX templates to personalized HTML emails
- Attaching relevant PDF documents
- Sending bulk emails via SMTP with rate limiting

**Impact**: Eliminated repetitive manual email tasks, allowing staff to focus on mission-critical work.

## Features

### Core Functionality
- **CSV-Driven Workflow**: Manages email campaigns through simple spreadsheet inputs
- **Document Conversion**: Converts `.docx` templates to both HTML and plain text for email compatibility
- **Multi-Attachment Support**: Automatically attaches personalized PDFs and organization documents
- **SMTP Integration**: Sends emails through configurable SMTP servers

### Safety & Testing
- **Dry-Run Mode**: Preview emails before sending (default behavior)
- **Debug Mode**: Test with hardcoded addresses before production run
- **Rate Limiting**: Built-in 1-second delay between sends to prevent throttling
- **Environment-Based Credentials**: Secure credential management via environment variables

## Usage

### Setup
```bash
# Install dependencies
pip install pandas mammoth

# Set SMTP credentials (only needed for actual sending)
export SMTP_USERNAME="your_username"
export SMTP_PASSWORD="your_password"
export SMTP_HOST="smtp.gmail.com"
export SMTP_PORT="587"
```

### CSV Format

Your CSV file should contain these columns:
```csv
DOCX,PDF_FILE,EMAIL,COMPANY
templates/letter1.docx,attachments/sponsor.pdf,contact@company.com,Company Name
```

### Running
```bash
# Preview mode (default - no emails sent)
python main.py --csv data.csv

# Send to test addresses (hardcoded in script)
python main.py --csv data.csv --send

# Send to actual recipients from CSV
python main.py --csv data.csv --send --actual
```

## Command-Line Arguments

| Flag | Description | Default |
|------|-------------|---------|
| `-c, --csv` | CSV file path (required) | - |
| `-s, --send` | Actually send emails | `False` (dry-run) |
| `-a, --actual` | Use CSV emails vs test emails | `False` (debug mode) |

## Architecture

### Key Design Decisions

**Separation of Concerns**: Email generation logic separated from SMTP sending via `EmailSender` callable class

**Progressive Testing**: Three-tier safety system (dry-run → test addresses → production)

**Modular Functions**: 
- `get_html_from_docx()` / `get_txt_from_docx()` - Document conversion
- `generate_email_message()` - Email composition
- `attach()` - File attachment handling
- `process()` - Main orchestration loop

**Configuration Over Code**: SMTP settings via environment variables, not hardcoded

## Technical Stack

- **Python 3.x**
- **pandas**: CSV parsing and data manipulation
- **mammoth**: DOCX to HTML/text conversion
- **smtplib**: Email sending via SMTP
- **email.message**: RFC-compliant email construction

## Example Workflow

1. **Prepare CSV** with recipient data (emails, company names, template paths)
2. **Create DOCX templates** with personalized content
3. **Test with dry-run**: `python main.py --csv data.csv`
4. **Verify with debug mode**: `python main.py --csv data.csv --send`
5. **Deploy to production**: `python main.py --csv data.csv --send --actual`

## Results

- ✅ **30+ hours saved annually** in manual email outreach
- ✅ **Personalized messaging** at scale for donor/sponsor campaigns
- ✅ **Zero production incidents** due to comprehensive testing modes
- ✅ **Extensible architecture** allowing easy addition of new campaigns

## Future Enhancements

- [ ] Template variable substitution (e.g., `{recipient_name}`)
- [ ] Email open/click tracking
- [ ] Retry logic for failed sends
- [ ] Campaign analytics dashboard
- [ ] Scheduled sending via cron/Task Scheduler

## License

Built for Chi-AWE (Chicago Asian Women Empowerment) nonprofit organization.

---

**Contact**: For questions about implementation or architecture decisions, feel free to reach out.
