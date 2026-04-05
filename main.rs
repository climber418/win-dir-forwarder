#![windows_subsystem = "windows"]

use log::{info, warn, error};
use log4rs;
use log4rs::{append::file::FileAppender, config::{Appender, Root}, encode::pattern::PatternEncoder, Config};

use windows::core::{Interface, Result, w};
use windows::Win32::{
    System::{
        Com::{
            CoCreateInstance, CoInitializeEx, CLSCTX_LOCAL_SERVER,
            COINIT_APARTMENTTHREADED, IServiceProvider,
        },
        Variant::VARIANT,
        SystemServices::SFGAO_FILESYSTEM,
    },
    UI::Shell::{
        IShellBrowser, IShellWindows, IShellView, IShellItem, IFolderView,
        ShellWindows, SIGDN_FILESYSPATH, SIGDN_DESKTOPABSOLUTEPARSING, SVGIO_SELECTION,
        IShellItemArray
    },
    UI::WindowsAndMessaging::{GetForegroundWindow, FindWindowExW, MessageBoxW, MB_ICONINFORMATION, MB_OK},
};
use windows::core::PCWSTR;
use windows::Win32::System::Com::IDispatch;

use std::env;
use std::process::Command;


unsafe fn get_selected_file_from_explorer() -> Result<String> {
	info!("get_selected_file_from_explorer:-->");

    info!("Initializing COM");
    let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);

    info!("Getting foreground window handle");
    let hwnd_gfw = GetForegroundWindow();
    info!("Foreground window handle: {:?}", hwnd_gfw);

    info!("Creating ShellWindows instance");
    let shell_windows: IShellWindows =
        CoCreateInstance(&ShellWindows, None, CLSCTX_LOCAL_SERVER)?;

    info!("Finding ShellTabWindowClass");
    let result_hwnd = match FindWindowExW(Some(hwnd_gfw), None, w!("ShellTabWindowClass"), None) {
        Ok(hwnd) => {
            info!("ShellTabWindowClass handle: {:?}", hwnd);
            hwnd
        }
        Err(e) => {
            info!("ShellTabWindowClass not found: {:?}, using foreground window", e);
            hwnd_gfw
        }
    };

    let mut target_path = String::new();
    let count = shell_windows.Count().unwrap_or_default();
    info!("Total shell windows count: {}", count);

    for i in 0..count {
        info!("Processing shell window index: {}", i);
        let variant = VARIANT::from(i);
        let dispatch: IDispatch = shell_windows.Item(&variant)?;

        let shell_browser = dispath2browser(dispatch);

        if shell_browser.is_none() {
            info!("Shell browser is None for index {}, skipping", i);
            continue;
        }
        let shell_browser = shell_browser.unwrap();
        info!("Successfully got shell browser for index {}", i);

        // 调用 GetWindow 可能会阻塞 GUI 消息
        let phwnd = shell_browser.GetWindow()?;
        info!("Shell browser window handle: {:?}", phwnd);

        if hwnd_gfw.0 == phwnd.0 || result_hwnd.0 == phwnd.0 {
            info!("Window handle matched for index {} (foreground: {:?}, target: {:?}, current: {:?})",
                  i, hwnd_gfw, result_hwnd, phwnd);
			let shell_view = shell_browser.QueryActiveShellView().unwrap();
            target_path = get_base_location_from_shellview(shell_view);
            info!("Retrieved base location: {}", target_path);
			break;
        } else {
            info!("Window handle mismatch for index {}, skipping (foreground: {:?}, target: {:?}, current: {:?})",
                  i, hwnd_gfw, result_hwnd, phwnd);
        }
    }
    info!("get_selected_file_from_explorer:<-- {}",target_path);
    Ok(target_path)
}

unsafe fn dispath2browser(dispatch: IDispatch) -> Option<IShellBrowser> {
    
    let mut service_provider: Option<IServiceProvider> = None;
    dispatch.query(&IServiceProvider::IID,
	    &mut service_provider as *mut _ as *mut _,
        )
        .ok()
        .unwrap();
    if service_provider.is_none() {
        return None;
    }
    let shell_browser = service_provider
        .unwrap()
        .QueryService::<IShellBrowser>(&IShellBrowser::IID)
        .ok();
    shell_browser
}

unsafe fn get_selected_file_path_from_shellview(shell_view: IShellView) -> String {
    let mut target_path = String::new();
    let shell_items = shell_view.GetItemObject::<IShellItemArray>(SVGIO_SELECTION);

    if shell_items.is_err() {
        return target_path;
    }
    info!("shell_items: {:?}", shell_items);
    let shell_items = shell_items.unwrap();
    let count = shell_items.GetCount().unwrap_or_default();
    for i in 0..count {
        let shell_item = shell_items.GetItemAt(i).unwrap();

        // 如果不是文件对象则继续循环
        if let Ok(attrs) = shell_item.GetAttributes(SFGAO_FILESYSTEM) {
            log::info!("attrs: {:?}", attrs);
            if attrs.0 == 0 {
                continue;
            }
        }

        if let Ok(display_name) = shell_item.GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING)
        {
            let tmp = display_name.to_string();
            if tmp.is_err() {
                continue;
            }
            target_path = tmp.unwrap();
            break;
        }

        if let Ok(display_name) = shell_item.GetDisplayName(SIGDN_FILESYSPATH) {
            let tmp = display_name.to_string();
            if tmp.is_err() {
                continue;
            }
            target_path = tmp.unwrap();
            break;
        }
        
    }
    target_path
}

unsafe fn get_base_location_from_shellview(shell_view: IShellView) -> String {
    let mut base_path = String::new();
    
    // Try to get the current folder from the shell view
    // We need to query for IFolderView interface to get folder information
    if let Ok(folder_view) = shell_view.cast::<windows::Win32::UI::Shell::IFolderView>() {
        if let Ok(folder) = folder_view.GetFolder::<IShellItem>() {
            // Try to get the file system path first
            if let Ok(display_name) = folder.GetDisplayName(SIGDN_FILESYSPATH) {
                if let Ok(path_str) = display_name.to_string() {
                    base_path = path_str;
                }
            }
            // Fallback to desktop absolute parsing name
            else if let Ok(display_name) = folder.GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING) {
                if let Ok(path_str) = display_name.to_string() {
                    base_path = path_str;
                }
            }
        }
    }
    
    base_path
}

fn main() -> Result<()> {

    // Initialize logging from the configuration file
    // log4rs::init_file("d:\\myproject\\win-dir-forwarder\\log4rs.yml", Default::default()).unwrap();

    // Create a custom JSON encoder
    let json_encoder = Box::new(PatternEncoder::new("{d} [{l}] - {m}{n}"));

    // Create a file appender with the custom encoder
    let file_appender = FileAppender::builder()
        .encoder(json_encoder)
        .build("d:\\myproject\\win-dir-forwarder\\logs\\log.txt")
        .unwrap();

    // Create a log configuration with the file appender
    let config = Config::builder()
        .appender(Appender::builder().build("file", Box::new(file_appender)))
        .build(Root::builder().appender("file").build(log::LevelFilter::Info))
        .unwrap();

    // Initialize the logger
    log4rs::init_config(config).unwrap();

    // Parse command line arguments
    let args: Vec<String> = env::args().collect();
    info!("Command line arguments:{}: {:?}",args.len(), args);

    // If no arguments, use the executable itself as target
    let target_exe = if args.len() >= 1 {
        args[1].clone()
    } else {
        info!("No target executable specified, using self as target");
        return Ok(());
    };

    let target_args = if args.len() >= 2 {
        &args[2..]
    } else {
        &[]
    };

    info!("Target executable: {}", &target_exe);
    info!("Target arguments: {:?}", target_args);

    // Get the current directory from Explorer
    let result = unsafe { get_selected_file_from_explorer() };
    match result {
        Ok(path) => {
            if path.is_empty() {
                warn!("Explorer directory path is empty, launching without setting working directory");
                // Launch without setting working directory
                match Command::new(&target_exe)
                    .args(target_args)
                    .spawn()
                {
                    Ok(child) => {
                        info!("Successfully launched {} with PID: {} (no working directory set)", &target_exe, child.id());
                    }
                    Err(e) => {
                        error!("Failed to launch {}: {:?}", &target_exe, e);
                    }
                }
            } else {
                info!("Working directory from Explorer: {}", path);

                // Launch the target executable with the specified working directory
                match Command::new(&target_exe)
                    .args(target_args)
                    .current_dir(&path)
                    .spawn()
                {
                    Ok(child) => {
                        info!("Successfully launched {} with PID: {} in directory: {}", &target_exe, child.id(), path);
                    }
                    Err(e) => {
                        error!("Failed to launch {} in directory {}: {:?}", &target_exe, path, e);
                    }
                }
            }
        }
        Err(e) => {
            error!("Error getting directory from Explorer: {:?}", e);
            warn!("Attempting to launch {} without setting working directory", &target_exe);
            // Try to launch anyway without setting working directory
            match Command::new(&target_exe)
                .args(target_args)
                .spawn()
            {
                Ok(child) => {
                    info!("Successfully launched {} with PID: {} (fallback mode)", &target_exe, child.id());
                }
                Err(e) => {
                    error!("Failed to launch {} in fallback mode: {:?}", &target_exe, e);
                }
            }
        }
    }

    Ok(())
}
