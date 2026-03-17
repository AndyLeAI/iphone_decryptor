# iPhone Decryptor

A desktop utility for reading and extracting data from local iPhone backups created with **Apple Devices** on Windows.

This app is built with **PySide6** and is designed for a simple workflow: create a local backup with Apple Devices, unlock the backup inside the app, then extract readable data such as messages, call history, contacts, photos, and voicemail.

## Features

- Auto-detect the latest iPhone backup on your computer
- Support for encrypted local backups
- Unlock backups with the backup password
- Choose a custom output folder
- Extract selected categories or everything at once
- Desktop GUI built with PySide6

## Important Notice

This tool is intended for:

- your own device backups, or
- backups you are explicitly authorized to access

Please use this project responsibly and in compliance with local laws, privacy rules, and device ownership requirements.

## Tech Stack

- Python
- PySide6
- Windows desktop environment
- Apple Devices backup workflow

## How to Use

1. Open **Microsoft Store**, search for **Apple Devices**, and install it.
2. Connect your iPhone to the PC with a USB cable.
3. In the **General** tab inside Apple Devices, choose **Back up all of the data on your iPhone to this computer**.
4. Turn on **Encrypt local backup** if you want to protect the backup with a password.
5. Click **Back Up Now** and wait for the backup to finish.
6. Open this app and click **Auto Find**. The app will locate the newest backup automatically.
7. If the backup is encrypted, enter the same password you used in Apple Devices and click **Unlock Backup**.
8. Choose your **Output Folder**. If you do not change it, the extracted files will be saved to your **Downloads** folder by default.
9. In **What to extract**, choose the categories you want to inspect, or select all if needed.
10. Click **Extract Selected** to start the extraction and decryption process.

## Build From Source

Clone the repository, install dependencies, and run the app:

```bash
pip install -r requirements.txt
python main.py
```

## Build Windows EXE

If you use the provided batch file:

```bat
BUILD.bat
```

This should build:

```text
dist\iPhone_decryptor.exe
```

## Suggested Repository Structure

```text
main.py
BUILD.bat
icons/
README.md
```

## License

This project is released under the **MIT License**.

MIT is a good fit here because it is simple, permissive, and easy for other developers to use, modify, and contribute to.

> Note: This project uses PySide6. Make sure you also review the license obligations of any third-party dependencies you distribute with your app.

## Contact

**Andy N Le**  
Email: me@andyle.one  
Facebook: facebook.com/ndle2

---

If you publish this on GitHub, you should also add a separate `LICENSE` file with the MIT license text.
