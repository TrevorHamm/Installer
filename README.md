EDS Installer: A simple installer that pulls down a zip file from a network 
location, extracts it and creates a shortcut to the extracted executable in 
the Start menu.

Usage: Installer.exe <program_name>

The zip file is extracted to the MyApps directory.
A shortcut is created in the Programs/MyApps Start menu.

Dependencies:
     ZLib:    https://github.com/kiyolee/zlib-win-build.git
     LibZip:  https://github.com/kiyolee/libzip-win-build.git

Application and dependencies are statically linked so that the executable
is a single file.

Author: Trevor Hamm

Actions:
- Get program name from commandline arguments
- Find newest zip file from network folder by that name
- Check / Install / Upgrade local installer   (STEP 1)
- Download zip to %localappdata%\MyApps       (STEP 2)
- Check/fail if program is currently running
- Uninstall current version (if exists)       (STEP 3)
- Unzip file                                  (STEP 4)
- Create shortcut                             (STEP 5)
- Run app on exit                             (STEP 6)

TODO:
- Add DEBUG flag (inconsistent results with what I have)
- Delete local copy of the application zip after extracting it.
     (This isn't working for some reason)
  
