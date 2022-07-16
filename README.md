# Maximilian's Automation Utility

Maximilian's automation utility is a Win32 in-process COM server library written in C++. It provides three easy-to-use utility services as COM automation objects: a job progress display, a dialog-based user input provider, and a program version resource reader.

Any client software capable of COM automation can receive the services. An application in C# can reference ProgressBox and immediately gain a job progress UI. A script program in JScript can use InputBox to gain a text entry dialog like a VB script can with VB's InputBox function. InputBox is more powerful than VBScript's. For one thing, it can show a check box if presenting an optional setting is needed. For another, it allows users to browse for a file or folder for its pathname as well. A script staging an installer build can use VersionInfo to effortlessly read an executable file's version number from its resource table. You just pass a pathname and read the VersionString property.



## Features

This is an automation programming library meant for Windows developers and IT professionals. Its main feature is MaxsUtilLib, a COM automation server written in C++. It provides the three automation objects of InputBox, ProgressBox, and VersionInfo. Use them in your script or C# application to quickly gain text input, progress output, and version query capabilities.

The library is accompanied with a couple of example programs. ListFileVersions is a WPF C# application. It uses MaxsUtilLib to list files in a folder, each with version info retrieved from the file's version resource. TestUtil is a console program written in C++. It programmatically tests the automation interfaces of MaxsUtilLib. Run this program to validate a MaxsUtilLib fresh out of the oven.

The distribution also includes a couple of JScript programs. One is TestMaxsUtil.js, a test script for verifying correct installation of the product. The other is stuffCab.js, a script for compressing a directory into a cab. Both demonstrate how the MaxsUtilLib automation objects can be incorporated to instantly gain practical user interface.


## Getting Started

Prerequisites:

* Windows 7 or newer. The library was originally written on Windows 10.
* Visual Studio 2017 or 2019. The VS solution and project files were generated using VS 2017. They can be imported by VS 2019.
* Windows SDK 10.0.17763.0. More recent SDKs should work, too, although no verification has been made.

MSI installers are available, if you don't want to be bothered with the source code and building the dll yourself. Install the library, and start incorporating the automation objects into your application.

* [x64 msi](https://github.com/mtanabe-sj/maximilians-automation-utility/blob/main/installer/out/MaxsUtil64.msi)
* [x86 msi](https://github.com/mtanabe-sj/maximilians-automation-utility/blob/main/installer/out/MaxsUtil86.msi)


## Using MaxsUtilLib

Refer to [api.docx](https://github.com/mtanabe-sj/maximilians-automation-utility/blob/main/api.docx). It provides a complete documentation of the automation API of MaxsUtilLib. It also contains programming tips with code samples to help you quickly achieve your automation needs.


## Contributing

Bug reports and enhancement requests are welcome.



## License

Copyright (c) mtanabe, All rights reserved.

The MaxsUtilLib and example applications are distributed under the terms of the MIT License.
