package testError

import (
	"errors"
	"fmt"
)

//TestError return (string, error)
func TestError(test string) (string, error) {
	if test == "" {
		return "", errors.New("empty input")
	}

	result := fmt.Sprintf("result = %v", test)
	return result, nil
}
