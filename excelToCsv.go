package main

import (
	"encoding/csv"
	"errors"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"io/ioutil"
	"os"
	"path"
	"path/filepath"
	"strings"
)

func ToCsvByFilename(fullFileName string, targetDir string) error {
	filenameWithSuffix := filepath.Base(fullFileName)
	fileSuffix := path.Ext(filenameWithSuffix)                         //获取文件后缀
	filenameOnly := strings.TrimSuffix(filenameWithSuffix, fileSuffix) //获取文件名
	if fileSuffix != ".xls" && fileSuffix != ".xlsx" {
		return errors.New("文件格式不正确")
	}

	originFile, err := excelize.OpenFile(fullFileName)
	if err != nil {
		fmt.Println(err)
		return err
	}

	sheetsMap := originFile.GetSheetMap()
	sheetsNum := len(sheetsMap)
	if sheetsNum > 1 {
		targetDir = targetDir + filenameOnly + "\\"
		_ = os.MkdirAll(targetDir, os.ModePerm)
	}

	for i, sheetName := range sheetsMap {
		targetFilename := filenameOnly + ".csv"
		// 多个sheet 把 根据文件名建立一个文件夹，放在文件夹下面
		if sheetsNum > 1 {
			targetFilename = fmt.Sprintf("Sheet%d_%s.csv", i, sheetName)
		}

		rows, err := originFile.GetRows(sheetName)
		if err != nil {
			fmt.Println(err)
			continue
		}

		//读取或创建目标文件
		file, err := os.OpenFile(targetDir+targetFilename, os.O_CREATE|os.O_RDWR, 0644)
		if err != nil {
			fmt.Println(err)
			continue
		}

		// 写入UTF-8 BOM，防止中文乱码
		_, _ = file.WriteString("\xEF\xBB\xBF")
		w := csv.NewWriter(file)
		for i, row := range rows { // 循环写入csv文件
			if i == 0 { // 写入头
				headSlice := make([]string, 0)
				for _, colCell := range row {
					headSlice = append(headSlice, colCell)
				}
				_ = w.Write(headSlice)
				w.Flush()
				continue
			}

			arr := make([]string, 0)
			for _, colCell := range row {
				arr = append(arr, colCell)
			}
			_ = w.Write(arr)

			// 写文件需要flush，不然缓存满了，后面的就写不进去了，只会写一部分 锘
			w.Flush()
		}
		file.Close()
	}

	return nil
}

// 文件后缀统一为\
func ToCsvByDir(originalDir string, targetDir string) error {
	// 路径处理
	if !strings.HasSuffix(originalDir, "/") && !strings.HasSuffix(originalDir, "\\") {
		originalDir = originalDir + "\\"
	}

	if strings.TrimSpace(targetDir) == "" {
		targetDir = originalDir + "csv\\"
	}

	if !strings.HasSuffix(targetDir, "/") && !strings.HasSuffix(targetDir, "\\") {
		targetDir = targetDir + "\\"
	}

	_, err := os.Stat(targetDir)
	if os.IsNotExist(err) {
		err = os.MkdirAll(targetDir, os.ModePerm)
		if err != nil {
			return err
		}
	}

	rd, err := ioutil.ReadDir(originalDir)
	if err != nil {
		fmt.Println(err)
		return err
	}
	// 路径处理结束
	//parent := getParentDirectory(originalDir)
	//fmt.Println(parent)
	for _, fi := range rd {
		if fi.IsDir() {
			if fi.Name() != "csv" {
				_ = ToCsvByDir(originalDir+fi.Name()+"\\", targetDir+fi.Name()+"\\")
			}
		} else {
			err = ToCsvByFilename(originalDir+fi.Name(), targetDir)
			if err != nil {
				fmt.Println(fmt.Errorf("file 【%s 】sync err:%s", fi.Name(), err))
			}
		}
	}
	return nil
}

func substr(s string, pos, length int) string {
	runes := []rune(s)
	l := pos + length
	if l > len(runes) {
		l = len(runes)
	}
	return string(runes[pos:l])
}

func getParentDirectory(dirctory string) string {
	if strings.LastIndex(dirctory, "\\") >= 0 {
		return substr(dirctory, 0, strings.LastIndex(dirctory, "\\"))
	}

	if strings.LastIndex(dirctory, "/") > 0 {
		return substr(dirctory, 0, strings.LastIndex(dirctory, "/"))
	}
	return dirctory
}
