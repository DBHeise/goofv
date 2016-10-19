# goofv
A command-line wrapper around Office File Validation (OFV)


Find more about OFV [here](https://support.microsoft.com/en-us/kb/2501584) and [here](https://blogs.technet.microsoft.com/office2010/2009/12/16/office-2010-file-validation/)

## Building From Source
assuming you already have GO installed and in the PATH
```batch
git clone https://github.com/DBHeise/goofv.git
cd goofv
set GOPATH=%CD%
go get
go build 
```

## Running goofv.exe
At this time, you MUST have Office installed on the same machine where goofv runs, it searches for the install location based on the installed version.

The command-line options are easily shown with the /? or -? argument:
```
Usage of goofv.exe:
  -file string
        file to validate
  -format string
        output format (must be one of: txt,csv,json,xml) (default "txt")
  -mode string
        validation mode (must be one of: xls,doc,ppt) (default "xls")
```

## Understanding failures
With the log settings that goofv forces, you will get an XML file in the %TMP% folder that contains details on the failure. It is an XML file that will probably have a bunch of information you don't care about, but IF you want to know the exact path inside the file where validation failed the GKError block in the XML is where you should start (you can use the typename and membername fields on the CallstackElement to search the
  [Binary Documentation](https://msdn.microsoft.com/en-us/library/office/cc313105(v=office.14).aspx)
, or in
  [OffVis](https://msdn.microsoft.com/en-us/library/office/gg615407(v=office.14).aspx)

## Difference between goofv and BFFValidator
BFFValidator is a tool that Microsoft sort-of released to validate binary files against the documentation and can be found [here](https://www.microsoft.com/en-us/download/details.aspx?id=26794). It is not the same validation that is run when a file is opened as the security requirements are different from the documentation requirements. For example, the documentation may say that a certian field MUST be (or SHOULD be) 0xVALUE, but in reality it doesn't matter what that value is because when the file is opened in the appropirate Office application, if it is too "big" or "small" it will just replace it with a default value

## Using goofv to detect malware
Many (but not all) known binary file exploits will be detected with by OFV, and can be categorized by their specific failure location.
