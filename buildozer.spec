[app]
title = Plankton Uploader
package.name = planktonuploader
package.domain = com.plankton
source.dir = .
source.include_exts = py,json
version = 1.0.0

# Application requirements
requirements = python3,kivy,pdfplumber,pdfminer.six,gspread,google-auth,google-auth-oauthlib,requests,charset-normalizer,certifi,urllib3,idna,cryptography,cffi,pyasn1,pyasn1-modules,rsa,cachetools,oauthlib,requests-oauthlib,pillow,plyer

# Permissions
android.permissions = INTERNET,READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE

# iOS settings
ios.kivy_ios_url = https://github.com/kivy/kivy-ios
ios.kivy_ios_branch = master
ios.ios_deploy_url = https://github.com/nicest/ios-deploy
ios.ios_deploy_branch = 1.10.0
ios.codesign.allowed = false

# Orientation
orientation = portrait

# Fullscreen
fullscreen = 0

# Icon (place icon.png in project root if desired)
# icon.filename = icon.png

# iOS-specific
[app:ios]
ios.permissions = NSPhotoLibraryUsageDescription,NSDocumentPickerUsageDescription

[buildozer]
log_level = 2
warn_on_root = 1
