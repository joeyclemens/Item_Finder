use actix_web::{web, App, HttpResponse, HttpServer, Responder, Result, Error, HttpRequest};
use askama::Template;
use serde::{Deserialize, Deserializer};
use tiberius::{AuthMethod, Client, Config};
use tokio_util::compat::TokioAsyncWriteCompatExt;
use futures_util::TryStreamExt;
use actix_web::http::header;
use futures_util::stream::Stream;
use std::pin::Pin;
use std::task::{Context, Poll};
use std::time::Duration;
use futures_util::stream;
use serde_json;
use rust_xlsxwriter::{Workbook, Format, XlsxError};
use actix_web::http::header::{CONTENT_TYPE, CONTENT_DISPOSITION};
use actix_web::error::ErrorInternalServerError;
use webbrowser;
use tokio::sync::oneshot;
use std::sync::Mutex;

#[derive(Template)]
#[template(path = "index.html")]
struct IndexTemplate;

#[derive(Deserialize)]
struct SearchQuery {
    terms: Vec<String>,
}

async fn index() -> impl Responder {
    let s = IndexTemplate.render().unwrap();
    HttpResponse::Ok().content_type("text/html").body(s)
}

async fn search(query: web::Json<SearchQuery>) -> Result<HttpResponse, Error> {
    let terms: Vec<String> = query.terms.iter().filter(|s| !s.trim().is_empty()).cloned().collect();
    if terms.is_empty() {
        return Ok(HttpResponse::BadRequest().body("No search terms provided"));
    }

    // Connect to SQL Server with Windows Auth
    let mut config = Config::new();
    config.host("mjm-sql01");
    config.trust_cert();
    config.authentication(AuthMethod::Integrated);
    config.port(1433);
    // No username/password needed for Integrated Auth

    let tcp = tokio::net::TcpStream::connect(("mjm-sql01", 1433)).await.map_err(|e| {
        actix_web::error::ErrorInternalServerError(format!("TCP connect error: {}", e))
    })?;
    let tcp = tcp.compat_write();
    let mut client = Client::connect(config.clone(), tcp).await.map_err(|e| {
        actix_web::error::ErrorInternalServerError(format!("DB connect error: {}", e))
    })?;

    // Get list of user databases (exclude system DBs)
    let dbs_query = "SELECT name FROM sys.databases WHERE database_id > 4";
    let mut dbs = Vec::new();
    {
        let mut stream = client.query(dbs_query, &[]).await.map_err(|e| {
            actix_web::error::ErrorInternalServerError(format!("DB list error: {}", e))
        })?;
        while let Some(item) = stream.try_next().await.map_err(|e| {
            actix_web::error::ErrorInternalServerError(format!("DB list row error: {}", e))
        })? {
            if let tiberius::QueryItem::Row(row) = item {
                let name: &str = row.get(0).unwrap();
                dbs.push(name.to_string());
            }
        }
    } // stream is dropped here, so client can be reused

    let mut results = Vec::new();
    for db in dbs {
        // Switch database
        let use_db = format!("USE [{}]", db);
        if let Err(_e) = client.simple_query(&use_db).await {
            continue; // skip DBs we can't use
        }
        // Build fuzzy search query
        let mut where_clauses = Vec::new();
        for term in &terms {
            where_clauses.push(format!("Item_Description LIKE '%{}%'", term.replace("'", "''")));
        }
        if where_clauses.is_empty() {
            continue; // skip this DB if no valid search terms
        }
        let where_sql = where_clauses.join(" OR ");
        let sql = format!(
            "SELECT '{}' as db, i.ADB_Ref, i.Item_Description, i.Unit_Cost, e.CAT, e.[Group] FROM Item_descriptions i LEFT JOIN ERM e ON e.ADB_Code = i.ADB_Ref WHERE {}",
            db, where_sql
        );
        if let Ok(mut stream) = client.query(&sql, &[]).await {
            while let Some(item) = stream.try_next().await.unwrap_or(None) {
                if let tiberius::QueryItem::Row(row) = item {
                    let db_name: Option<&str> = row.get(0);
                    let adb_ref: Option<&str> = row.get(1);
                    let desc: Option<&str> = row.get(2);
                    let unit_cost: Option<f64> = row.get(3);
                    let cat: Option<&str> = row.get(4);
                    
                    // Handle Group column which might be float or string
                    let group_str = match row.try_get::<f64, _>(5) {
                        Ok(Some(g)) => {
                            if g.fract() == 0.0 {
                                format!("[Group: {}]", g as i64)
                            } else {
                                format!("[Group: {:.2}]", g)
                            }
                        },
                        Ok(None) => String::from("[Group: ]"),
                        Err(_) => {
                            // If float conversion fails, try as string
                            match row.try_get::<&str, _>(5) {
                                Ok(Some(g_str)) => {
                                    if let Ok(g) = g_str.parse::<f64>() {
                                        if g.fract() == 0.0 {
                                            format!("[Group: {}]", g as i64)
                                        } else {
                                            format!("[Group: {:.2}]", g)
                                        }
                                    } else {
                                        format!("[Group: {}]", g_str)
                                    }
                                },
                                Ok(None) => String::from("[Group: ]"),
                                Err(_) => String::from("[Group: ]"), // Fallback if both fail
                            }
                        }
                    };

                    // Find which terms matched (case-insensitive)
                    let mut matched_terms = Vec::new();
                    let desc_lower = desc.unwrap_or("").to_lowercase();
                    for term in &terms {
                        if desc_lower.contains(&term.to_lowercase()) {
                            matched_terms.push(term);
                        }
                    }
                    let matched_str = format!(" (matched: {})", terms.join(", "));
                    let unit_cost_str = match unit_cost {
                        Some(cost) => format!("[Unit_Cost: £{:.2}]", cost),
                        None => String::from("[Unit_Cost: ]"),
                    };
                    let cat_str = cat.unwrap_or("");
                    
                    results.push(format!(
                        "[{}] {} - {}{} {} [CAT: {}] {}",
                        db_name.unwrap_or(""),
                        adb_ref.unwrap_or(""),
                        desc.unwrap_or(""),
                        matched_str,
                        unit_cost_str,
                        cat_str,
                        group_str
                    ));
                }
            }
        }
    }
    let body = if results.is_empty() {
        "No results found".to_string()
    } else {
        results.join("\n")
    };
    Ok(HttpResponse::Ok().content_type("text/plain").body(body))
}

async fn export_excel(query: web::Json<SearchQuery>) -> Result<HttpResponse, Error> {
    let terms: Vec<String> = query.terms.iter().filter(|s| !s.trim().is_empty()).cloned().collect();
    if terms.is_empty() {
        return Ok(HttpResponse::BadRequest().body("No search terms provided"));
    }

    // Connect to SQL Server with Windows Auth
    let mut config = Config::new();
    config.host("mjm-sql01");
    config.trust_cert();
    config.authentication(AuthMethod::Integrated);
    config.port(1433);

    let tcp = tokio::net::TcpStream::connect(("mjm-sql01", 1433)).await.map_err(|e| {
        actix_web::error::ErrorInternalServerError(format!("TCP connect error: {}", e))
    })?;
    let tcp = tcp.compat_write();
    let mut client = Client::connect(config.clone(), tcp).await.map_err(|e| {
        actix_web::error::ErrorInternalServerError(format!("DB connect error: {}", e))
    })?;

    // Get list of user databases (exclude system DBs)
    let dbs_query = "SELECT name FROM sys.databases WHERE database_id > 4";
    let mut dbs = Vec::new();
    {
        let mut stream = client.query(dbs_query, &[]).await.map_err(|e| {
            actix_web::error::ErrorInternalServerError(format!("DB list error: {}", e))
        })?;
        while let Some(item) = stream.try_next().await.map_err(|e| {
            actix_web::error::ErrorInternalServerError(format!("DB list row error: {}", e))
        })? {
            if let tiberius::QueryItem::Row(row) = item {
                let name: &str = row.get(0).unwrap();
                dbs.push(name.to_string());
            }
        }
    }

    // Prepare Excel workbook in memory
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    // Set column widths (in Excel's character units)
    worksheet.set_column_width(0, 34.37).map_err(ErrorInternalServerError)?;
    worksheet.set_column_width(1, 13.58).map_err(ErrorInternalServerError)?;
    worksheet.set_column_width(2, 90.11).map_err(ErrorInternalServerError)?;
    worksheet.set_column_width(3, 18.32).map_err(ErrorInternalServerError)?;
    worksheet.set_column_width(4, 14.0).map_err(ErrorInternalServerError)?;
    worksheet.set_column_width(5, 14.0).map_err(ErrorInternalServerError)?;
    worksheet.set_column_width(6, 14.0).map_err(ErrorInternalServerError)?;
    // Freeze the top row
    worksheet.set_freeze_panes(1, 0).map_err(ErrorInternalServerError)?;
    // Add General Sans fonts to the workbook
    let header_format = Format::new()
        .set_background_color(0xF5F5F5)
        .set_font_name("General Sans Medium");
    let wrap_format = Format::new()
        .set_text_wrap()
        .set_font_name("General Sans");
    worksheet.write_string_with_format(0, 0, "Database", &header_format).map_err(ErrorInternalServerError)?;
    worksheet.write_string_with_format(0, 1, "ADB_Ref", &header_format).map_err(ErrorInternalServerError)?;
    worksheet.write_string_with_format(0, 2, "Description", &header_format).map_err(ErrorInternalServerError)?;
    worksheet.write_string_with_format(0, 3, "Matched Terms", &header_format).map_err(ErrorInternalServerError)?;
    worksheet.write_string_with_format(0, 4, "Unit_Cost", &header_format).map_err(ErrorInternalServerError)?;
    worksheet.write_string_with_format(0, 5, "CAT", &header_format).map_err(ErrorInternalServerError)?;
    worksheet.write_string_with_format(0, 6, "Group", &header_format).map_err(ErrorInternalServerError)?;

    let mut row_idx = 1;
    for db in dbs {
        let use_db = format!("USE [{}]", db);
        if let Err(_e) = client.simple_query(&use_db).await {
            continue;
        }
        let mut where_clauses = Vec::new();
        for term in &terms {
            where_clauses.push(format!("Item_Description LIKE '%{}%'", term.replace("'", "''")));
        }
        if where_clauses.is_empty() {
            continue;
        }
        let where_sql = where_clauses.join(" OR ");
        let sql = format!(
            "SELECT '{}' as db, i.ADB_Ref, i.Item_Description, i.Unit_Cost, e.CAT, e.[Group] FROM Item_descriptions i LEFT JOIN ERM e ON e.ADB_Code = i.ADB_Ref WHERE {}",
            db, where_sql
        );
        if let Ok(mut stream) = client.query(&sql, &[]).await {
            while let Some(item) = stream.try_next().await.unwrap_or(None) {
                if let tiberius::QueryItem::Row(row) = item {
                    let db_name: Option<&str> = row.get(0);
                    let adb_ref: Option<&str> = row.get(1);
                    let desc: Option<&str> = row.get(2);
                    let unit_cost: Option<f64> = row.get(3);
                    let cat: Option<&str> = row.get(4);
                    
                    let mut matched_terms = Vec::new();
                    let desc_lower = desc.unwrap_or("").to_lowercase();
                    for term in &terms {
                        if desc_lower.contains(&term.to_lowercase()) {
                            matched_terms.push(term.as_str());
                        }
                    }
                    
                    worksheet.write_string_with_format(row_idx, 0, db_name.unwrap_or(""), &wrap_format).map_err(ErrorInternalServerError)?;
                    worksheet.write_string_with_format(row_idx, 1, adb_ref.unwrap_or(""), &wrap_format).map_err(ErrorInternalServerError)?;
                    worksheet.write_string_with_format(row_idx, 2, desc.unwrap_or(""), &wrap_format).map_err(ErrorInternalServerError)?;
                    worksheet.write_string_with_format(row_idx, 3, &matched_terms.join(", "), &wrap_format).map_err(ErrorInternalServerError)?;
                    if let Some(cost) = unit_cost {
                        worksheet.write_string_with_format(row_idx, 4, &format!("£{:.2}", cost), &wrap_format).map_err(ErrorInternalServerError)?;
                    } else {
                        worksheet.write_string_with_format(row_idx, 4, "", &wrap_format).map_err(ErrorInternalServerError)?;
                    }
                    worksheet.write_string_with_format(row_idx, 5, cat.unwrap_or(""), &wrap_format).map_err(ErrorInternalServerError)?;
                    
                    // Handle Group column which might be float or string
                    match row.try_get::<f64, _>(5) {
                        Ok(Some(g)) => {
                            if g.fract() == 0.0 {
                                worksheet.write_string_with_format(row_idx, 6, &format!("{}", g as i64), &wrap_format).map_err(ErrorInternalServerError)?;
                            } else {
                                worksheet.write_string_with_format(row_idx, 6, &format!("{:.2}", g), &wrap_format).map_err(ErrorInternalServerError)?;
                            }
                        },
                        Ok(None) => {
                            worksheet.write_string_with_format(row_idx, 6, "", &wrap_format).map_err(ErrorInternalServerError)?;
                        },
                        Err(_) => {
                            // If float conversion fails, try as string
                            match row.try_get::<&str, _>(5) {
                                Ok(Some(g_str)) => {
                                    if let Ok(g) = g_str.parse::<f64>() {
                                        if g.fract() == 0.0 {
                                            worksheet.write_string_with_format(row_idx, 6, &format!("{}", g as i64), &wrap_format).map_err(ErrorInternalServerError)?;
                                        } else {
                                            worksheet.write_string_with_format(row_idx, 6, &format!("{:.2}", g), &wrap_format).map_err(ErrorInternalServerError)?;
                                        }
                                    } else {
                                        worksheet.write_string_with_format(row_idx, 6, g_str, &wrap_format).map_err(ErrorInternalServerError)?;
                                    }
                                },
                                Ok(None) => {
                                    worksheet.write_string_with_format(row_idx, 6, "", &wrap_format).map_err(ErrorInternalServerError)?;
                                },
                                Err(_) => {
                                    worksheet.write_string_with_format(row_idx, 6, "", &wrap_format).map_err(ErrorInternalServerError)?;
                                }
                            }
                        }
                    }
                    
                    row_idx += 1;
                }
            }
        }
    }
    let buffer = workbook.save_to_buffer().map_err(ErrorInternalServerError)?;
    Ok(HttpResponse::Ok()
        .append_header((CONTENT_TYPE, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
        .append_header((CONTENT_DISPOSITION, "attachment; filename=Item_Finder_Results.xlsx"))
        .body(buffer))
}

async fn shutdown(tx: web::Data<Mutex<Option<oneshot::Sender<()>>>>) -> impl Responder {
    if let Some(tx) = tx.lock().unwrap().take() {
        let _ = tx.send(()); // trigger shutdown
        HttpResponse::Ok().body("Server is shutting down...")
    } else {
        HttpResponse::Ok().body("Shutdown already triggered.")
    }
}

#[actix_web::main]
async fn main() -> std::io::Result<()> {
    // Spawn a thread to open the browser after a short delay
    std::thread::spawn(|| {
        std::thread::sleep(std::time::Duration::from_millis(500));
        let _ = webbrowser::open("http://127.0.0.1:8082");
    });

    let (tx, rx) = oneshot::channel::<()>();
    let tx_data = web::Data::new(Mutex::new(Some(tx)));

    let server = HttpServer::new(move || {
        App::new()
            .app_data(tx_data.clone())
            .route("/", web::get().to(index))
            .route("/search", web::post().to(search))
            .route("/export_excel", web::post().to(export_excel))
            .route("/shutdown", web::post().to(shutdown))
    })
    .bind(("127.0.0.1", 8082))?
    .run();

    tokio::select! {
        _ = server => {},
        _ = rx => {},
    }

    Ok(())
}