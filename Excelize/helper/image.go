package helper

import (
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"strings"
)

func DownloadImage(url string) string {

	// don't worry about errors
	response, e := http.Get(url)
	name := strings.Split(url, "signatures/")
	if e != nil {
		log.Fatal(e)
	}
	defer response.Body.Close()

	fileName := name[len(name)-1] + ".jpeg"
	//open a file for writing
	file, err := os.Create(fileName)
	if err != nil {
		log.Fatal(err)
	}

	defer file.Close()

	// Use io.Copy to just dump the response body to the file. This supports huge files
	_, err = io.Copy(file, response.Body)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("Success!")
	return fileName
}

func RemoveImage(name string) {

	err := os.Remove(name)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("Remove Success!")
}
