use std::{
    env, fs,
    io::{self, Write},
    path::PathBuf,
    thread,
    time::{Duration, SystemTime},
};
use wincent::{
    handle::{add_to_recent_files, remove_from_recent_files},
    query::is_in_recent_files,
    WincentResult,
    error::WincentError,
};

struct ScopedFile {
    path: PathBuf,
}

impl ScopedFile {
    fn new() -> io::Result<Self> {
        let timestamp = SystemTime::now()
            .duration_since(SystemTime::UNIX_EPOCH)
            .unwrap()
            .as_millis();
        
        let filename = format!("wincent-test-{}.txt", timestamp);
        let path = env::current_dir()?.join(filename);

        let mut file = fs::File::create(&path)?;
        writeln!(file, "This is a test file for Quick Access operations")?;

        Ok(Self { path })
    }

    fn path_str(&self) -> Option<&str> {
        self.path.to_str()
    }
}

impl Drop for ScopedFile {
    fn drop(&mut self) {
        if let Err(e) = fs::remove_file(&self.path) {
            eprintln!("Failed to delete temporary file: {}", e);
        }
    }
}

fn main() -> WincentResult<()> {
    let file = ScopedFile::new()?;
    let file_path = file.path_str().ok_or_else(|| WincentError::InvalidPath("Invalid file path encoding".to_string()))?;

    println!("Working with temporary file: {}", file_path);

    // Add file to recent items
    println!("Adding file to Quick Access...");
    add_to_recent_files(file_path)?;

    // Wait for Windows to update
    thread::sleep(Duration::from_millis(1000));

    // Verify if file has been added
    if is_in_recent_files(file_path)? {
        println!("File successfully added to Quick Access");
    } else {
        println!("Failed to add file to Quick Access");
        return Ok(());
    }

    // Remove file from recent items
    println!("Removing file from Quick Access...");
    remove_from_recent_files(file_path)?;

    // Wait for Windows to update
    thread::sleep(Duration::from_millis(500));

    // Verify if file has been removed
    if !is_in_recent_files(file_path)? {
        println!("File successfully removed from Quick Access");
    } else {
        println!("Failed to remove file from Quick Access");
    }

    // Temporary file will be automatically deleted when temp_file goes out of scope
    Ok(())
}
