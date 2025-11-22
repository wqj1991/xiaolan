// Learn more about Tauri commands at https://tauri.app/develop/calling-rust/
use std::collections::HashMap;
use std::fs;
use std::path::Path;

#[tauri::command]
fn greet(name: &str) -> String {
    format!("Hello, {}! You've been greeted from Rust!", name)
}

#[tauri::command]
async fn select_and_scan_word_folder() -> HashMap<String, serde_json::Value> {
    let mut result = HashMap::new();
    
    // 在Tauri v2中使用对话框API
    if let Some(folder_path) = rfd::AsyncFileDialog::new()
        .set_title("选择Word文件所在文件夹")
        .set_directory(std::env::current_dir().unwrap_or_default())
        .pick_folder()
        .await
    {
        let folder_path_buf = folder_path.path();
        let folder_path_str = folder_path_buf.to_string_lossy().to_string();
        
        // 扫描Word文件
        match scan_word_files(&folder_path_buf) {
            Ok(word_files) => {
                // 构建返回结果
                result.insert("folder_path".to_string(), serde_json::Value::String(folder_path_str));
                
                // 转换文件列表为JSON数组
                let files_json: Vec<_> = word_files
                    .iter()
                    .map(|(name, path)| {
                        let mut file_obj = HashMap::new();
                        file_obj.insert("name".to_string(), serde_json::Value::String(name.clone()));
                        file_obj.insert("path".to_string(), serde_json::Value::String(path.clone()));
                        serde_json::Value::Object(file_obj.into_iter().collect())
                    })
                    .collect();
                
                result.insert("word_files".to_string(), serde_json::Value::Array(files_json));
            },
            Err(err) => {
                result.insert("error".to_string(), serde_json::Value::String(err.to_string()));
            }
        }
    } else {
        // 用户取消选择
        result.insert("error".to_string(), serde_json::Value::String("用户取消选择".to_string()));
    }
    
    result
}

// 扫描文件夹中的Word文件
fn scan_word_files(folder_path: &Path) -> Result<Vec<(String, String)>, String> {
    let mut word_files = Vec::new();
    
    match fs::read_dir(folder_path) {
        Ok(entries) => {
            for entry in entries {
                match entry {
                    Ok(entry) => {
                        let path = entry.path();
                        
                        if path.is_file() {
                            if let Some(file_name) = path.file_name().and_then(|f| f.to_str()) {
                                // 检查是否为Word文件
                                if file_name.to_lowercase().ends_with(".docx") || 
                                   file_name.to_lowercase().ends_with(".doc") {
                                    word_files.push((
                                        file_name.to_string(),
                                        path.to_string_lossy().to_string()
                                    ));
                                }
                            }
                        }
                    },
                    Err(err) => {
                        return Err(format!("读取文件项失败: {}", err));
                    }
                }
            }
            
            // 按文件名排序
            word_files.sort_by(|a, b| a.0.cmp(&b.0));
            
            Ok(word_files)
        },
        Err(err) => {
            Err(format!("读取文件夹失败: {}", err))
        }
    }
}

// 重命名Word文件
#[tauri::command]
async fn rename_word_file(old_path: String, new_name: String) -> Result<String, String> {
    let old_path_buf = Path::new(&old_path);
    
    // 获取文件所在目录
    let parent_dir = match old_path_buf.parent() {
        Some(dir) => dir,
        None => return Err("无法获取文件目录".to_string()),
    };
    
    // 构建新文件路径
    let new_path = parent_dir.join(&new_name);
    
    // 检查新文件名是否已存在
    if new_path.exists() {
        return Err(format!("文件 \"{}\" 已存在", new_name));
    }
    
    // 执行重命名
    match fs::rename(old_path_buf, &new_path) {
        Ok(_) => Ok(new_path.to_string_lossy().to_string()),
        Err(err) => Err(format!("重命名失败: {}", err)),
    }
}

// 复制文件到指定文件夹（使用新文件名）
#[tauri::command]
async fn copy_file_to_folder(file_path: String, target_folder: String, new_name: String) -> Result<String, String> {
    let file_path_buf = Path::new(&file_path);
    let target_folder_buf = Path::new(&target_folder);
    
    // 检查文件是否存在
    if !file_path_buf.exists() {
        return Err(format!("文件不存在: {}", file_path));
    }
    
    // 创建目标文件夹（如果不存在）
    if let Err(err) = fs::create_dir_all(&target_folder_buf) {
        return Err(format!("创建文件夹失败: {}", err));
    }
    
    // 构建目标文件路径
    let target_path = target_folder_buf.join(&new_name);
    
    // 检查目标文件是否已存在
    if target_path.exists() {
        // 如果已存在，添加时间戳后缀
        let stem = Path::new(&new_name)
            .file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("file");
        let extension = Path::new(&new_name)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("");
        
        let timestamp = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_secs();
        
        let final_name = if extension.is_empty() {
            format!("{}_{}", stem, timestamp)
        } else {
            format!("{}_{}.{}", stem, timestamp, extension)
        };
        
        let target_path = target_folder_buf.join(&final_name);
        
        match fs::copy(file_path_buf, &target_path) {
            Ok(_) => Ok(target_path.to_string_lossy().to_string()),
            Err(err) => Err(format!("复制文件失败: {}", err)),
        }
    } else {
        // 执行复制
        match fs::copy(file_path_buf, &target_path) {
            Ok(_) => Ok(target_path.to_string_lossy().to_string()),
            Err(err) => Err(format!("复制文件失败: {}", err)),
        }
    }
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![greet, select_and_scan_word_folder, rename_word_file, copy_file_to_folder])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
