# Item Finder

A Rust-based web application for searching and exporting item descriptions from multiple SQL Server databases. Results can be viewed in the browser or exported to Excel.

## Features
- Search for item descriptions across all user databases on a SQL Server instance
- View results in the browser
- Export search results to an Excel file
- Uses Actix Web for the backend and Askama for HTML templating

## Prerequisites
- Rust (latest stable recommended)
- SQL Server instance accessible on your network (configured for Windows Authentication)
- [General Sans](https://github.com/GeneralSans/GeneralSans) font files (included in `font/`)
- Windows OS (for integrated authentication)

## Build & Run

1. **Clone the repository**
2. **Install dependencies:**
   ```sh
   cargo build
   ```
3. **Run the application:**
   ```sh
   cargo run
   ```
   This will start a web server at [http://127.0.0.1:8082](http://127.0.0.1:8082) and open it in your browser.

## Usage
- Enter search terms in the web interface and submit to search.
- Click the export button to download results as an Excel file.
- To shut down the server, close the terminal window or use the `/shutdown` endpoint.

## Configuration
- The SQL Server host is currently set to `mjm-sql01` in `src/main.rs`. Change this if your server is different.
- Only user databases (not system DBs) are searched.

## License
MIT 