# goofv
A command-line wrapper around Office File Validation


Find more about Office File Validation [here](https://support.microsoft.com/en-us/kb/2501584) and [here](https://blogs.technet.microsoft.com/office2010/2009/12/16/office-2010-file-validation/)

## Building From Source
assuming you already have GO installed and in the PATH
```batch
    git clone https://github.com/DBHeise/goofv.git
    cd goofv
    set GOPATH=%CD%
    go get -v golang.org/x/sys/windows/registry
    go build 
```
