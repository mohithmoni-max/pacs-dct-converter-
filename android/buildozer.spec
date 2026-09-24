[app]
title = DCT Mobile Validator
package.name = dctmobilevalidator
package.domain = com.mohithmoni
source.dir = .
source.include_exts = py,json
version = 1.0.0
orientation = portrait
fullscreen = 0
android.api = 35
android.minapi = 26
android.ndk_api = 26
android.archs = arm64-v8a
android.accept_sdk_license = True
requirements = hostpython3==3.12.14,python3==3.12.14,kivy==2.3.1,numpy,pandas,openpyxl,xlrd,androidstorage4kivy
p4a.branch = master
android.debug_artifact = apk
log_level = 2

[buildozer]
warn_on_root = 0
