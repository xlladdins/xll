# xll - a library for creating Excel add-ins

This library is a C++ wrapper for the Microsoft
[Excel Software Development Kit](https://docs.microsoft.com/en-us/office/client-developer/excel/welcome-to-the-excel-software-development-kit)
that makes it simple to create add-ins that register C/C++ functions 
that can be called from Excel.

## Prerequisites

<dl>
<dt>Windows 10</dt>
<dd>The Excel SDK is not supported on MacOS.</dd>
</dl>


[Visual Studio 2019](https://visualstudio.microsoft.com/)  
  Use the Community Edition and choose the `Desktop development with C++` workload.</dd>


<dl>
<dt>[Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel)</dt>
<dd>Install the 64-bit version of Office 365 for the best experience.</dd>
</dl>

<dl>
<dt>[Sandcastle Help File Builder](https://github.com/EWSoftware/SHFB)</dt>
<dd>This is optional if you want to create help files integrated into Excel.</dd>
</dl>

## Get Started

Fork the [xll](https://github.com/xlladdins/xll) repository and
[clone](https://docs.microsoft.com/en-us/visualstudio/get-started/tutorial-open-project-from-repo)
it in Visual Studio.
Run the `setup.bat` script in the top level `xll` folder.

Create a new project using `File ► New ► Project...` (`Ctrl-Shift-N`) and
select `XLL Project`. At this point you can compile and run the add-in
using `Debug ► Start Debuggiing` (`F5`). 
