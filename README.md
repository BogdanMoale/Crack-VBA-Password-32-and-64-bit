# VBA Password Cracker

This project presents a method for cracking VBA (Visual Basic for Applications) passwords in Excel workbooks. VBA passwords are used to protect VBA code from being viewed or edited without authorization.

## Overview

The VBA code provided in this project utilizes dynamic memory manipulation techniques to bypass the password protection on VBA projects within Excel workbooks. By intercepting calls to certain Windows API functions related to VBA project access, this code can effectively remove or bypass the password protection mechanism.

## Features

- **Dynamic Hooking**: The code dynamically hooks into specific Windows API functions used in VBA project access to intercept and modify their behavior.
- **Password Bypass**: By intercepting calls related to VBA project password verification, this code effectively bypasses the need for a password to access or modify VBA code.
- **Excel Workbook Compatibility**: This code is designed to work specifically with Excel workbooks containing VBA projects protected by passwords.

## Usage

To utilize this functionality:
1. Import the provided VBA code into your Excel workbook.
2. Ensure macros are enabled in your Excel settings.
3. Open the workbook containing the password-protected VBA project.
4. Run the provided code to remove or bypass the VBA project password protection.

## Considerations

- **Legal and Ethical Implications**: Cracking passwords without proper authorization may violate terms of service, user agreements, or even laws depending on jurisdiction. Ensure that you have the legal right to access and modify the VBA code.
- **Backup Original Files**: Before attempting to crack VBA passwords, it's recommended to create backups of the original Excel workbooks to avoid data loss or corruption.
- **Use with Caution**: Manipulating password protection mechanisms can have unintended consequences and may lead to instability or loss of functionality in Excel workbooks.

## Credits

This project is inspired by the need for developers and users to access locked VBA projects for legitimate purposes such as recovering lost passwords or modifying legacy code.
