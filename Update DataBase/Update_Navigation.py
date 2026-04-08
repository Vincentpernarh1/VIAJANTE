import dotenv
import os
from playwright.sync_api import sync_playwright
from datetime import datetime
import shutil

# Load environment variables using absolute path
script_dir = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(script_dir, "..", "BD", ".env")
dotenv.load_dotenv(dotenv_path=env_path)
sharepoint_url = os.getenv("SHAREPOINT_URL")

if not sharepoint_url:
    raise ValueError(f"SHAREPOINT_URL not found in .env file at {env_path}")

print(f"SharePoint URL: {sharepoint_url}")

# Create the download folder path with date
current_date = datetime.now().strftime("%Y-%m-%d")
download_folder = os.path.abspath(rf"..\BD")

# Ensure the download folder exists
os.makedirs(download_folder, exist_ok=True)

print(f"Download folder: {download_folder}")

# Use AppData for automation profile (standard for .exe applications)
# Each user will have their own isolated profile with their own SSO login
automation_profile = os.path.join(os.getenv('LOCALAPPDATA'), 'Viajante', 'edge_automation_profile')
os.makedirs(automation_profile, exist_ok=True)

# Files to download
FILES_TO_DOWNLOAD = [
    "BD_CADASTRO_PN",
    "BD_CADASTRO_MDR"
]

def cleanup_old_versions(filename, current_file, silent=False):
    """Delete old versions of the file, keeping only the newly downloaded one"""
    import glob
    import re
    
    # Find all files in the download folder
    all_files = os.listdir(download_folder)
    
    # Pattern to match: {filename}_YYYY-MM-DD.extension (dated backups only)
    date_pattern = re.compile(rf"^{re.escape(filename)}_\d{{4}}-\d{{2}}-\d{{2}}\..+$")
    
    deleted_count = 0
    for file in all_files:
        # Check if it matches our dated backup pattern
        if date_pattern.match(file):
            file_path = os.path.join(download_folder, file)
            # Skip the file we just downloaded
            if os.path.abspath(file_path) != os.path.abspath(current_file):
                try:
                    os.remove(file_path)
                    if not silent:
                        print(f"  🗑️  Deleted old version: {file}")
                    deleted_count += 1
                except Exception as e:
                    if not silent:
                        print(f"  ⚠️  Could not delete {file}: {e}")
    
    if deleted_count == 0 and not silent:
        print(f"  ℹ️  No old versions to clean up")
    
    return deleted_count

def download_file_from_sharepoint(page, filename, silent=False):
    """Download a specific file from the SharePoint page"""
    if not silent:
        print(f"\nLooking for {filename} file...")
    
    # Wait a moment for page to be ready
    page.wait_for_timeout(2000)
    
    # Try to find and click the file
    file_found = False
    
    # Try different approaches to find the file
    selectors_to_try = [
        f"button[name*='{filename}']",
        f"a[title*='{filename}']",
        f"[aria-label*='{filename}']",
        f"text={filename}.xlsx",
        f"text={filename}",
    ]
    
    for selector in selectors_to_try:
        try:
            element = page.locator(selector).first
            if element.is_visible(timeout=2000):
                if not silent:
                    print(f"  ✓ Found file using selector: {selector}")
                
                # Right-click to open context menu
                element.click(button="right")
                page.wait_for_timeout(1000)
                
                # Look for "Download" option in context menu
                download_option = page.locator("text=Download").first
                
                # Click download and wait for the download to start
                with page.expect_download(timeout=30000) as download_info:
                    download_option.click()
                
                download = download_info.value
                file_found = True
                
                # Save the file with a timestamp
                original_filename = download.suggested_filename
                file_extension = original_filename.split('.')[-1] if '.' in original_filename else 'xlsx'
                saved_filename = f"{filename}_{current_date}.{file_extension}"
                save_path = os.path.join(download_folder, saved_filename)
                download.save_as(save_path)
                
                if not silent:
                    print(f"  ✓ Downloaded successfully to: {saved_filename}")
                
                # Clean up old versions
                cleanup_old_versions(filename, save_path, silent=silent)
                
                return True
        except Exception as e:
            continue
    
    if not file_found:
        if not silent:
            print(f"  ✗ Could not find {filename} file")
        return False
    
    return file_found

def download_sharepoint_files(headless=False, silent=False, auto_close=False):
    """Main function to download all required files from SharePoint
    
    Args:
        headless: If True, run browser in headless mode (no window)
        silent: If True, suppress print messages
        auto_close: If True, close browser automatically without waiting for input
    """
    with sync_playwright() as p:
        # Launch Edge with automation profile
        if not silent:
            print("Launching Edge browser for automation...")
        context = p.chromium.launch_persistent_context(
            user_data_dir=automation_profile,
            headless=headless,
            channel="msedge",
            accept_downloads=True,
            args=["--start-maximized"] if not headless else []
        )
        
        page = context.pages[0] if context.pages else context.new_page()
        
        try:
            # Navigate to SharePoint
            if not silent:
                print(f"Navigating to SharePoint...")
                print(f"\nIMPORTANT: On first run, you'll need to:")
                print("1. Login with your Microsoft account (SSO)")
                print("2. The session will be saved for future runs")
                print()
            
            page.goto(sharepoint_url, wait_until="domcontentloaded", timeout=60000)
            if not silent:
                print("Page loaded!")
            
            # Wait for the page to fully load
            page.wait_for_timeout(3000)
            
            # Download each file
            results = {}
            for filename in FILES_TO_DOWNLOAD:
                success = download_file_from_sharepoint(page, filename, silent=silent)
                results[filename] = success
            
            # Summary
            if not silent:
                print("\n" + "="*70)
                print("Download Summary:")
                print("="*70)
                for filename, success in results.items():
                    status = "✓ SUCCESS" if success else "✗ FAILED"
                    print(f"{status}: {filename}")
            
            return results  # Return results for programmatic use
            
        except Exception as e:
            if not silent:
                print(f"\n✗ Error: {e}")
                print("Taking a screenshot for debugging...")
            page.screenshot(path="../debug_screenshot.png")
            if not silent:
                print("Screenshot saved to: debug_screenshot.png")
            return {f: False for f in FILES_TO_DOWNLOAD}
        
        finally:
            if not silent and not auto_close:
                input("\nPress Enter to close the browser...")
            context.close()

if __name__ == "__main__":
    print("="*70)
    print("SharePoint Files Downloader")
    print("="*70)
    print(f"Files to download: {', '.join(FILES_TO_DOWNLOAD)}")
    print(f"Profile location: {automation_profile}")
    print()
    download_sharepoint_files()
    print("\nDone!")
