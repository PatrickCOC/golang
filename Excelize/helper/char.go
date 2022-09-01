package helper

func ToChar(i int) string {
	abc := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	return abc[i-1 : i]
}

func ToInt(r []rune) int {
	return int(r[0]) - 64
}

// func ToChar(i int) string {
// 	var foo = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
// 	return string(foo[i-1])
// }
