# goofv
A command-line wrapper around Office File Validation


Find more about Office File Validation [here](https://support.microsoft.com/en-us/kb/2501584) and [here](https://blogs.technet.microsoft.com/office2010/2009/12/16/office-2010-file-validation/)

## Building From Source
assuming you already have GO installed and in the PATH
```batch
    git clone https://github.com/DBHeise/goofv.git
    cd goofv
    set GOPATH=%CD%
    go get
    go build 
```

## Understanding failures
With the log settings that goofv forces, you will get an XML file in the %TMP% folder that contains details on the failure. It is an XML file that will probably have a bunch of information you don't care about, but IF you want to know the exact path inside the file where validation failed the GKError block in the XML is where you should start (you can use the typename and membername fields on the CallstackElement to search the
  [Binary Documentation](https://msdn.microsoft.com/en-us/library/office/cc313105(v=office.14).aspx)
, or in
  [OffVis](https://msdn.microsoft.com/en-us/library/office/gg615407(v=office.14).aspx)