package main

import (
	"flag"
	"fmt"
)

var (
	dir    = flag.String("dir", "", "original dir")
	target = flag.String("target", "", "target dir")
)

func main() {
	flag.Parse()

	if *dir == "" {
		fmt.Println(fmt.Sprintf("help:\n\t dir is require"))
		return
	}

	//"D:\\source\\eternity_configurations\\share\\data"
	err := ToCsvByDir(*dir, *target)
	if err != nil {
		fmt.Println(err)
	} else {
		fmt.Println("转换完成！！！")
	}
}
