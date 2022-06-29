package main

import (
	"fmt"

	testError "example/test/function"
)

func main() {
	fmt.Println("testing")
	testError.TestError("I am test")
}
