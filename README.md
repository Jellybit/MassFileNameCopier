# Mass File Name Copier
A PowerShell script with a drag-n-drop GUI to simplify and speed up renaming multiple files by copying names from one set to another, or creating a bunch of file names in some other editor.

<img width="750" height="563" alt="image" src="https://github.com/user-attachments/assets/9e1c1146-50f3-4c7b-90d8-90879df5e8cb" />

I often need to rename a bunch of files to match the names of another set, like when organizing media files, or just to create my own naming style in Excel and mass rename based on that. Batch renaming tools don’t let you directly copy names from one set to another, and they also tend to have issues with non-English characters, or other special characters. So, I created this drag-n-drop GUI tool to make the process fast for me, especially for complex file names with special characters. Most of the time, I just get the file names, and make all the changes I want in Notepad++, then send it back to the original files. But I also like making all new file names in Excel because I can just drag down cells to cover episode numbers, and other things like that.

## Features

* **Flexible Renaming:**
  - Drag or select a folder or multiple files to rename.
  - Extract file names from one set and apply them to another in the same order.
  - Or just work with one set of files and make all the changes you want.
  - Edit names directly in a text box for custom renaming.
  - Or paste text that you've created elsewhere.
* **Smart Path Input:**
  - **Browse Button:** Manually select a folder or files.
  - **Drag & Drop:** Drag files or a folder into the input path box, or drag files into the names box to extract their names.
* **Safety First:**
  - Checks if the number of new names matches the number of files to rename.
  - Validates paths and skips invalid or non-existent files.
  - Shows clear error messages for issues like empty paths or mismatched counts.
* **Unsafety Second:**
  - I don't confirm a rename, to speed things up for myself, so be careful and double check your work.

### Steps to use:
- **Input Path:**
  - Drag a folder (e.g., `D:\Videos`) or multiple files (e.g., shift/ctrl select and drag `file1.mkv`, `file2.mkv`) into the input path box, or click **Browse** to select them.
  - The path will show as the folder path or file paths separated by |||.
- **Get File Names:**
  - Click **Get File Names** to list the names (without extensions) of the files in the input path.
  - Alternatively, drag another set of files into the file names box to extract their names in one action, or type/paste custom names (one per line).
- **Rename Files:**
  - Edit the names in the file names box if needed.
  - Click **Rename Files** to apply the names in the text box to the files in the input path, keeping their original extensions.
  - You’ll get a confirmation message when renaming is successful, or an error if something goes wrong (e.g., mismatched counts).

## Important Notes

* **Extensions Not Considered:** Since I've had to copy file names in groups of files of various video formats, I decided not to make it concern itself with renaming the extensions. That way, I could have a big list of MKV files, and make it use the names of old and mixed avi, mp4, etc... files without changing the claimed format of the new files.
* **Permissions:** The script requires write access to the files’ directory. If you encounter permission errors, try running PowerShell as Administrator.
* **PowerShell Version:** Requires PowerShell 5.1 or later (included with Windows 10/11). Theoretically. I only use Powershell 7.

## Compile Standalone

To create a standalone executable without the console window, follow these steps in PowerShell:

```powershell
Install-Module ps2exe
Set-ExecutionPolicy RemoteSigned -Scope Process -Force
ps2exe "MassFileNameCopier.ps1" "MassFileNameCopier.exe" -noConsole
```

This installs `ps2exe`, which compiles PowerShell scripts. The execution policy command temporarily allows compilation, reverting when you close the PowerShell window.

---

Hope this tool helps someone. I use it all the time.
