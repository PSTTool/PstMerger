# Free PST Merge Tool for Outlook

**The ultimate open-source utility to merge, combine, and join multiple Outlook PST files into a single master file.**  
*Completely free, no size limits, and enterprise-grade reliability.*

![License](https://img.shields.io/badge/license-MIT-green.svg)
![Platform](https://img.shields.io/badge/platform-Windows_10%20%7C%2011%20%7C%20Server-blue.svg)
![Outlook](https://img.shields.io/badge/Outlook-2013%20%7C%202016%20%7C%202019%20%7C%202021%20%7C%20365-orange.svg)
![.NET](https://img.shields.io/badge/.NET-4.5+-purple.svg)

## Overview

**PST Merge Tool** is a powerful Windows application designed to consolidate unlimited `.pst` files into one managed archive. Unlike paid alternatives, this tool is 100% free and open source.

It was built specifically to handle **large mailboxes (up to 2TB)** and includes a registry fix to bypass the default 50GB Outlook file size limit. Perfect for IT administrators performing backup recovery or users organizing years of emails.

## Features

- ✅ **Merge unlimited PST files** into one master PST
- ✅ **2TB size limit support** (removes the default 50GB Outlook restriction)
- ✅ **Preserves folder structure** (Inbox, Sent Items, subfolders, etc.)
- ✅ **Handles all item types** (Emails, Contacts, Calendar, Tasks, Notes)
- ✅ **Graceful cancellation** - stop safely without corrupting data
- ✅ **Persistent logging** - every action saved to a log file
- ✅ **Disk space validation** - warns before running out of space
- ✅ **Error resilience** - continues processing even if individual items fail

## Requirements

- Windows 10/11 or Windows Server
- Microsoft Outlook installed (uses Outlook Interop)
- .NET Framework 4.5 or later

## Usage

1. **Download** the latest release from [Releases](../../releases)
2. **Run** `PstMerger.exe`
3. **Click "Fix PST Size Limits"** to remove the 50GB restriction (one-time setup)
4. **Select source folder** containing your PST files
5. **Select destination** for the merged PST
6. **Click "Start Merge"** and wait for completion

> ⚠️ **Important**: Close Outlook before running the merge. Do not use the computer for heavy tasks during the merge process.

## Screenshots

![PST Merge Tool Interface](screenshots/pst-merge-tool-ui.png)

## Support & Feedback

**Found a bug?**  
This project is open source and we welcome your feedback!  
Please [Submit an Issue](https://github.com/mithundgnxt-stack/PstMerger/issues) if you encounter any problems or have feature ideas.

## Building from Source

```bash
# Clone the repository
git clone https://github.com/YourUsername/PstMergeTool.git
cd PstMergeTool

# Build with MSBuild
msbuild PstMerger.csproj /p:Configuration=Release
```

The compiled executable will be in `bin\Release\PstMerger.exe`.

## How It Works

The tool uses Microsoft Outlook Interop to:
1. Open each source PST file
2. Recursively traverse all folders
3. Copy items to the destination PST, preserving folder structure
4. Safely close and detach each PST after processing

This ensures maximum compatibility and data integrity, as Outlook handles all the low-level PST format details.

## Registry Fix Details

By default, Outlook limits PST files to ~50GB. The "Fix PST Size Limits" button modifies:

```
HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Outlook\PST
HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\PST
```

Setting `MaxLargeFileSize` and `WarnLargeFileSize` to ~2TB.

## License

MIT License - see [LICENSE](LICENSE) file.

## Author

**Mithun** 

© 2026 All Rights Reserved.
