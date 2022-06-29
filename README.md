# golang
## Download & Install

* Run
    sudo apt update
    sudo apt upgrade 

* Search for Go:
    sudo apt search golang-go
    sudo apt search gccgo-go

* Install
    sudo apt install golang-go

* golang Vesion
    go version

## Create Project

    $ go mod init example/hello
    go: creating new go.mod: module example/hello

### Main.go

    package main

    func main() {
        #code here
    }

### Run

    go run .

### Add bew module requirements

    go mod tidy