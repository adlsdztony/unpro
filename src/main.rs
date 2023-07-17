// Import the necessary dependencies
use regex::Regex;
use std::fs::File;
use std::io::prelude::*;
use std::path::Path;
use zip::ZipArchive;

fn decompress_xlsx(file: &str) -> Result<String, Box<dyn std::error::Error>> {
    // Get directory
    let dir = Path::new(&file).file_stem().unwrap().to_str().unwrap();

    // Create directory
    std::fs::create_dir_all(&dir)?;

    // Decompress
    let mut zip = ZipArchive::new(File::open(file)?)?;
    for i in 0..zip.len() {
        let mut file = zip.by_index(i)?;
        let outpath = match file.enclosed_name() {
            Some(path) => path.to_owned(),
            None => continue,
        };
        let outpath = Path::new(&dir).join(outpath);

        if (&*file.name()).ends_with('/') {
            std::fs::create_dir_all(&outpath)?;
        } else {
            if let Some(p) = outpath.parent() {
                if !p.exists() {
                    std::fs::create_dir_all(&p)?;
                }
            }
            let mut outfile = File::create(&outpath)?;
            std::io::copy(&mut file, &mut outfile)?;
        }
    }

    println!("Decompressed {}", file);

    Ok(dir.to_string())
}

fn unlock_xlsx_workbook(dir: &str) -> Result<(), Box<dyn std::error::Error>> {
    let workbook_path = String::from(dir) + "/xl/workbook.xml";

    let mut content = String::new();
    let mut file = File::open(&workbook_path)?;
    file.read_to_string(&mut content)?;

    // Check if protected
    if content.contains("workbookProtection") {
        // Remove protection
        let re = Regex::new(r#"<workbookProtection.*?/>"#)?;
        content = re.replace_all(&content, "").to_string();
        println!("Removed protection from {:?}", &workbook_path);
    }

    // Check if hidden sheets
    if content.contains(r#"state="hidden""#) {
        // Remove hidden sheets
        let re = Regex::new(r#"state="hidden" "#)?;
        content = re.replace_all(&content, "").to_string();
        println!("Removed hidden sheets from {:?}", &workbook_path);
    }

    // Write content back to file
    let mut file = File::create(&workbook_path)?;
    file.write_all(content.as_bytes())?;

    Ok(())
}

fn get_files_recursive(dir: &str) -> Result<Vec<String>, Box<dyn std::error::Error>> {
    let mut files = Vec::new();
    let walker = walkdir::WalkDir::new(dir).follow_links(true);

    for entry in walker {
        let entry = entry?;
        if entry.file_type().is_file() {
            let entry_path = entry.path().display().to_string();
            // println!("Found {}", &entry_path);
            files.push(entry_path);
        }
    }

    Ok(files)
}

// get files in same level only, not recursive
fn get_files(dir: &str) -> Result<Vec<String>, Box<dyn std::error::Error>> {
    let mut files = Vec::new();
    let paths = std::fs::read_dir(dir).unwrap();
    for path in paths {
        let path = path.unwrap().path();
        if path.is_file() {
            // println!("Found {}", path.display());
            files.push(path.display().to_string());
        }
    }

    Ok(files)
}

fn rm_protection(path: &str) -> Result<(), Box<dyn std::error::Error>> {
    let mut content = String::new();
    let mut file = File::open(path)?;
    file.read_to_string(&mut content)?;

    // Check if protected
    if content.contains("sheetProtection") {
        // Remove protection
        let re = Regex::new(r#"<sheetProtection.*?/>"#)?;
        content = re.replace_all(&content, "").to_string();
        // Write content back to file
        let mut file = File::create(path)?;
        file.write_all(content.as_bytes())?;
        println!("Removed protection from {}", path);
    }

    Ok(())
}

fn unlock_xlsx_worksheets(dir_path: &str) -> Result<(), Box<dyn std::error::Error>> {
    let worksheets_path = String::from(dir_path) + "/xl/worksheets";
    let files = get_files_recursive(&worksheets_path)?;

    for file in files {
        if file.ends_with(".xml") {
            rm_protection(&file)?;
        }
    }

    Ok(())
}

fn compress_xlsx(dir: &str) -> Result<(), Box<dyn std::error::Error>> {
    let files = get_files_recursive(&dir)?;
    let zip_path = format!("{}_unpro.xlsx", &dir);

    let file = File::create(&zip_path)?;
    let mut zip = zip::ZipWriter::new(file);

    for file_path in files {
        let entry_path = Path::new(&file_path).strip_prefix(&dir)?;
        zip.start_file(
            entry_path.display().to_string(),
            zip::write::FileOptions::default(),
        )?;
        let mut file = File::open(&file_path)?;
        std::io::copy(&mut file, &mut zip)?;
    }

    println!("Compressed {}", zip_path);

    Ok(())
}

fn cleanup(dir: &str) -> Result<(), Box<dyn std::error::Error>> {
    std::fs::remove_dir_all(dir)?;
    println!("Removed {}", dir);

    Ok(())
}

fn unpro_xlsx(file: &str) -> Result<(), Box<dyn std::error::Error>> {
    let dir = decompress_xlsx(file)?;
    unlock_xlsx_workbook(&dir)?;
    unlock_xlsx_worksheets(&dir)?;
    compress_xlsx(&dir)?;
    cleanup(&dir)?;
    println!("Successfully unprotected {}", file);
    Ok(())
}

fn auto() -> Result<(), Box<dyn std::error::Error>> {
    // Get xlsx file from corrent directory
    let files = get_files(".")?;

    // rm non-xlsx files
    let files = files
        .into_iter()
        .filter(|f| f.ends_with(".xlsx"))
        .collect::<Vec<String>>();

    // check if any xlsx files found
    if files.len() == 0 {
        println!("No xlsx files found");

        // pause
        println!("Press enter to exit...");
        let mut input = String::new();
        std::io::stdin().read_line(&mut input)?;

        return Ok(());
    }

    // ask for confirmation
    for f in &files {
        println!("{}", f);
    }
    println!(
        "Found {} xlsx files. Continue? (any key to continue, n to abort)",
        files.len()
    );
    let mut input = String::new();
    std::io::stdin().read_line(&mut input)?;
    if input.trim() == "n" {
        println!("Aborted");
        return Ok(());
    }

    // Unprotect each xlsx file
    for f in files {
        unpro_xlsx(&f)?;
    }

    Ok(())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = std::env::args().collect();

    if args.len() == 1 {
        auto()?;
    } else if args.len() >= 2 {
        for file in &args[1..] {
            // check if file exists
            if !Path::new(file).exists() {
                println!("File not found: {}", file);
                continue;
            }

            // unprotect xlsx file
            unpro_xlsx(file)?;
        }
    } else {
        println!("Usage: unpro_xlsx [file1] [file2] ...");
    }

    // pause
    println!("Press enter to exit...");
    let mut input = String::new();
    std::io::stdin().read_line(&mut input)?;

    Ok(())
}
