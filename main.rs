#![windows_subsystem = "windows"]

use log::{info, warn, error};
use log4rs;
use log4rs::{append::file::FileAppender, config::{Appender, Root}, encode::pattern::PatternEncoder, Config};

use windows::core::{Interface, Result};
use windows::Win32::{
    Foundation::{BOOL, HWND, LPARAM},
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
        IShellItemArray, IShellFolderViewDual, Folder2, IWebBrowser2,
    },
    UI::WindowsAndMessaging::{
        GetForegroundWindow, MessageBoxW, MB_ICONINFORMATION, MB_OK,
        EnumChildWindows, EnumWindows, GetAncestor, GetClassNameW, GetWindowTextW,
        IsIconic, IsWindowVisible, GA_ROOTOWNER,
    },
};
use windows::core::PCWSTR;
use windows::Win32::System::Com::IDispatch;

use std::env;
use std::process::Command;
use std::path::Path;


/// 标签页窗口类名：Windows 11 中每个标签页都是一个该类的子窗口
const TAB_CLASS: &str = "ShellTabWindowClass";

#[derive(Default)]
struct TabSearch {
    active: Option<HWND>,
    first: Option<HWND>,
}

unsafe extern "system" fn find_tab_cb(hwnd: HWND, lparam: LPARAM) -> BOOL {
    let search = &mut *(lparam.0 as *mut TabSearch);

    let mut class_name = [0u16; 64];
    let len = GetClassNameW(hwnd, &mut class_name);
    if len <= 0 || String::from_utf16_lossy(&class_name[..len as usize]) != TAB_CLASS {
        return true.into();
    }

    if IsWindowVisible(hwnd).as_bool() {
        // 可见的标签页就是当前激活的标签页，无需继续枚举
        search.active = Some(hwnd);
        return false.into();
    }
    if search.first.is_none() {
        search.first = Some(hwnd);
    }
    true.into()
}

/// 找出框架窗口下当前激活的标签页窗口
unsafe fn active_tab(frame: HWND) -> Option<HWND> {
    let mut search = TabSearch::default();
    let _ = EnumChildWindows(
        Some(frame),
        Some(find_tab_cb),
        LPARAM(&mut search as *mut _ as isize),
    );
    search.active.or(search.first)
}

/// 读取 Shell 窗口项所属的框架窗口句柄（等价于 AHK 的 window.hwnd）
unsafe fn frame_of(dispatch: &IDispatch) -> Option<HWND> {
    let web_browser: IWebBrowser2 = dispatch.cast().ok()?;
    let hwnd = web_browser.HWND().ok()?.0;
    if hwnd == 0 {
        None
    } else {
        Some(HWND(hwnd as *mut core::ffi::c_void))
    }
}

struct ShellWindow {
    dispatch: IDispatch,
    browser: IShellBrowser,
    frame: Option<HWND>,
    tab: HWND,
}

struct TopmostSearch<'a> {
    windows: &'a [ShellWindow],
    found: Option<usize>,
}

/// EnumWindows 按 Z 序枚举（最近激活的在前），取第一个可见的候选窗口
unsafe extern "system" fn topmost_cb(hwnd: HWND, lparam: LPARAM) -> BOOL {
    let search = &mut *(lparam.0 as *mut TopmostSearch);
    if !IsWindowVisible(hwnd).as_bool() || IsIconic(hwnd).as_bool() {
        return true.into();
    }
    let matched = search
        .windows
        .iter()
        .position(|w| w.frame.map_or(false, |f| f.0 == hwnd.0));
    if let Some(index) = matched {
        search.found = Some(index);
        return false.into();
    }
    true.into()
}

unsafe fn window_desc(hwnd: HWND) -> String {
    let mut class_name = [0u16; 128];
    let class_len = GetClassNameW(hwnd, &mut class_name).max(0) as usize;
    let mut title = [0u16; 256];
    let title_len = GetWindowTextW(hwnd, &mut title).max(0) as usize;
    format!(
        "{} / {}",
        String::from_utf16_lossy(&class_name[..class_len]),
        String::from_utf16_lossy(&title[..title_len])
    )
}

unsafe fn collect_windows(shell_windows: &IShellWindows) -> Vec<ShellWindow> {
    let count = shell_windows.Count().unwrap_or_default();
    info!("Total shell windows count: {}", count);

    let mut windows = Vec::new();
    for i in 0..count {
        let variant = VARIANT::from(i);
        let dispatch: IDispatch = match shell_windows.Item(&variant) {
            Ok(dispatch) => dispatch,
            Err(e) => {
                warn!("Failed to get shell window index {}: {:?}", i, e);
                continue;
            }
        };

        let browser = match to_browser(&dispatch) {
            Some(browser) => browser,
            None => {
                info!("Shell browser is None for index {}, skipping", i);
                continue;
            }
        };

        let tab = match browser.GetWindow() {
            Ok(hwnd) => hwnd,
            Err(e) => {
                warn!("GetWindow failed for index {}: {:?}", i, e);
                continue;
            }
        };
        let frame = frame_of(&dispatch);

        info!("Shell window {}: frame={:?}, tab={:?}", i, frame, tab);
        windows.push(ShellWindow {
            dispatch,
            browser,
            frame,
            tab,
        });
    }
    windows
}

/// 同一框架窗口下有多个标签页时，挑出当前激活的那个
unsafe fn pick_tab<'a>(windows: &'a [ShellWindow], frame: HWND) -> Option<&'a ShellWindow> {
    let on_frame = |w: &&ShellWindow| w.frame.map_or(false, |f| f.0 == frame.0);

    if let Some(tab) = active_tab(frame) {
        if let Some(target) = windows.iter().find(|w| on_frame(w) && w.tab.0 == tab.0) {
            return Some(target);
        }
    }
    windows.iter().find(on_frame)
}

/// 找到前台窗口对应的浏览对象。
/// AHK 用 window.hwnd（框架窗口）和前台窗口比较，而 IShellBrowser::GetWindow
/// 返回的是标签页窗口，所以两者都参与匹配；多个标签页共享框架窗口时再按激活标签页筛选。
/// 如果前台窗口不是 Explorer（例如经由启动器、命令窗口触发），退回到 Z 序最靠前的
/// 可见 Explorer 窗口，也就是最近使用过的那个窗口。
unsafe fn find_browser(
    shell_windows: &IShellWindows,
    fg: HWND,
) -> Option<(IDispatch, IShellBrowser)> {
    let windows = collect_windows(shell_windows);
    if windows.is_empty() {
        warn!("No shell window candidate found");
        return None;
    }

    // 1. 前台窗口本身就是 Shell 窗口，或者是 Shell 窗口的弹出窗口
    let root_owner = GetAncestor(fg, GA_ROOTOWNER);
    for window in &windows {
        if let Some(frame) = window.frame {
            if frame.0 == fg.0 || frame.0 == root_owner.0 {
                let target = pick_tab(&windows, frame)?;
                return Some((target.dispatch.clone(), target.browser.clone()));
            }
        }
        if window.tab.0 == fg.0 {
            return Some((window.dispatch.clone(), window.browser.clone()));
        }
    }

    // 2. 前台窗口不是 Explorer，取最近激活的 Explorer 窗口
    let mut search = TopmostSearch {
        windows: &windows,
        found: None,
    };
    let _ = EnumWindows(
        Some(topmost_cb),
        LPARAM(&mut search as *mut _ as isize),
    );

    let index = match search.found {
        Some(index) => index,
        None => {
            warn!("All shell windows are hidden or minimized");
            return None;
        }
    };
    let window = &windows[index];
    let frame = window.frame?;
    info!(
        "Foreground window is not explorer, using most recently used frame: {:?}",
        frame
    );
    let target = pick_tab(&windows, frame).unwrap_or(window);
    Some((target.dispatch.clone(), target.browser.clone()))
}

/// 取窗口所属的文档对象，即 AHK 里的 window.document
unsafe fn document_view(dispatch: &IDispatch) -> Option<IShellFolderViewDual> {
    if let Ok(view) = dispatch.cast::<IShellFolderViewDual>() {
        return Some(view);
    }
    let web_browser: IWebBrowser2 = dispatch.cast().ok()?;
    web_browser.Document().ok()?.cast().ok()
}

/// 等价于 AHK 的 window.document.Folder.Self.Path
unsafe fn folder_path(dispatch: &IDispatch) -> String {
    let view = match document_view(dispatch) {
        Some(view) => view,
        None => {
            warn!("IShellFolderViewDual is not available for this shell window");
            return String::new();
        }
    };
    let folder = match view.Folder() {
        Ok(folder) => folder,
        Err(e) => {
            warn!("Failed to get folder from shell view: {:?}", e);
            return String::new();
        }
    };
    // Self 属性定义在 Folder2 上
    let item = match folder.cast::<Folder2>().and_then(|folder2| folder2.Self_()) {
        Ok(item) => item,
        Err(e) => {
            warn!("Failed to get folder self item: {:?}", e);
            return String::new();
        }
    };
    match item.Path() {
        Ok(path) => path.to_string(),
        Err(e) => {
            warn!("Failed to get folder path: {:?}", e);
            String::new()
        }
    }
}

unsafe fn explorer_path() -> Result<String> {
    info!("explorer_path:-->");

    info!("Initializing COM");
    let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);

    info!("Getting foreground window handle");
    let fg = GetForegroundWindow();
    info!(
        "Foreground window handle: {:?} ({})",
        fg,
        window_desc(fg)
    );
    if fg.0.is_null() {
        warn!("No foreground window, cannot resolve explorer directory");
        return Ok(String::new());
    }

    info!("Creating ShellWindows instance");
    let shell_windows: IShellWindows =
        CoCreateInstance(&ShellWindows, None, CLSCTX_LOCAL_SERVER)?;

    let mut path = String::new();
    match find_browser(&shell_windows, fg) {
        Some((dispatch, browser)) => {
            path = folder_path(&dispatch);
            if path.is_empty() {
                warn!("Folder self path is empty, falling back to shell view folder");
                if let Ok(shell_view) = browser.QueryActiveShellView() {
                    path = view_folder_path(shell_view);
                }
            }
        }
        None => warn!("No shell window matches the foreground window"),
    }

    info!("explorer_path:<-- {}", path);
    Ok(path)
}

unsafe fn to_browser(dispatch: &IDispatch) -> Option<IShellBrowser> {
    let service_provider: IServiceProvider = dispatch.cast().ok()?;
    service_provider
        .QueryService::<IShellBrowser>(&IShellBrowser::IID)
        .ok()
}

unsafe fn selected_file_path(shell_view: IShellView) -> String {
    let mut path = String::new();
    let shell_items = shell_view.GetItemObject::<IShellItemArray>(SVGIO_SELECTION);

    if shell_items.is_err() {
        return path;
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
            path = tmp.unwrap();
            break;
        }

        if let Ok(display_name) = shell_item.GetDisplayName(SIGDN_FILESYSPATH) {
            let tmp = display_name.to_string();
            if tmp.is_err() {
                continue;
            }
            path = tmp.unwrap();
            break;
        }
        
    }
    path
}

unsafe fn view_folder_path(shell_view: IShellView) -> String {
    let mut path = String::new();
    
    // Try to get the current folder from the shell view
    // We need to query for IFolderView interface to get folder information
    if let Ok(folder_view) = shell_view.cast::<windows::Win32::UI::Shell::IFolderView>() {
        if let Ok(folder) = folder_view.GetFolder::<IShellItem>() {
            // Try to get the file system path first
            if let Ok(display_name) = folder.GetDisplayName(SIGDN_FILESYSPATH) {
                if let Ok(path_str) = display_name.to_string() {
                    path = path_str;
                }
            }
            // Fallback to desktop absolute parsing name
            else if let Ok(display_name) = folder.GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING) {
                if let Ok(path_str) = display_name.to_string() {
                    path = path_str;
                }
            }
        }
    }
    
    path
}

// Check if the executable is MSYS2 or Git Bash that needs CHERE_INVOKING
fn is_msys_or_gitbash(exe_path: &str) -> bool {
    let exe_lower = exe_path.to_lowercase();
    let file_name = Path::new(&exe_lower)
        .file_name()
        .and_then(|n| n.to_str())
        .unwrap_or("");

    matches!(file_name,
        "clang64.exe" |
        "clangarm64.exe" |
        "mingw64.exe" |
        "msys.exe" |
        "ucrt64.exe"
    )
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
    let result = unsafe { explorer_path() };
    match result {
        Ok(path) => {
            let needs_msys_env = is_msys_or_gitbash(&target_exe);

            if path.is_empty() {
                warn!("Explorer directory path is empty, launching without setting working directory");
                // Launch without setting working directory
                let mut cmd = Command::new(&target_exe);
                cmd.args(target_args);

                if needs_msys_env {
                    cmd.env("CHERE_INVOKING", "1")
                       .env("MSYS2_PATH_TYPE", "inherit");
                }

                match cmd.spawn() {
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
                // For MSYS2/Git Bash, set CHERE_INVOKING=1 to prevent cd to home directory
                let mut cmd = Command::new(&target_exe);
                cmd.args(target_args).current_dir(&path);

                if needs_msys_env {
                    info!("Detected MSYS2/Git Bash shell, setting CHERE_INVOKING environment");
                    cmd.env("CHERE_INVOKING", "1")
                       .env("MSYS2_PATH_TYPE", "inherit");
                }

                match cmd.spawn() {
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

            let needs_msys_env = is_msys_or_gitbash(&target_exe);

            // Try to launch anyway without setting working directory
            let mut cmd = Command::new(&target_exe);
            cmd.args(target_args);

            if needs_msys_env {
                cmd.env("CHERE_INVOKING", "1")
                   .env("MSYS2_PATH_TYPE", "inherit");
            }

            match cmd.spawn() {
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
