// Import the necessary dependencies
use std::fs::File;
use std::io::prelude::*;
use std::path::Path;
use std::env;
use zip::ZipArchive;
use regex::Regex;


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


fn get_files(dir: &str) -> Result<Vec<String>, Box<dyn std::error::Error>> {
    let mut files = Vec::new();
    let walker = walkdir::WalkDir::new(dir).follow_links(true);

    for entry in walker {
        let entry = entry?;
        if entry.file_type().is_file() {
            let entry_path = entry.path().display().to_string();
            files.push(entry_path);
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
    let files = get_files(&worksheets_path)?;

    for file in files {
        if file.ends_with(".xml") {
            rm_protection(&file)?;
        }
    }

    Ok(())
}

fn compress_xlsx(dir: &str) -> Result<(), Box<dyn std::error::Error>> {
    let files = get_files(&dir)?;
    let zip_path = format!("{}_unpro.xlsx", &dir);
    
    let file = File::create(&zip_path)?;
    let mut zip = zip::ZipWriter::new(file);

    for file_path in files {
        let entry_path = Path::new(&file_path).strip_prefix(&dir)?;
        zip.start_file(entry_path.display().to_string(), zip::write::FileOptions::default())?;
        let mut file = File::open(&file_path)?;
        std::io::copy(&mut file, &mut zip)?;
    }

    println!("Compressed {}", zip_path);

    Ok(())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Get command-line argument
    let args: Vec<String> = env::args().collect();

    // Check if there is an argument
    if args.len() < 2 {
        println!("Usage: unpro file.xlsx");
        std::process::exit(1);
    }

    let file = &args[1];

    // Decompress
    let dir = decompress_xlsx(file)?;

    // Unlock workbook
    unlock_xlsx_workbook(&dir)?;

    // Unlock worksheets
    unlock_xlsx_worksheets(&dir)?;

    // Compress
    compress_xlsx(&dir)?;

    // Remove directory
    std::fs::remove_dir_all(&dir)?;

    Ok(())
}
