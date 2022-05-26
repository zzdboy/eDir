package main

import (
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/storage"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

func init() {
	os.Setenv("FYNE_THEME", "light")
}

func main() {

	//新建一个app
	a := app.New()

	//设置程序图标
	a.SetIcon(resourceIcoPng)

	// 中文字体显示
	a.Settings().SetTheme(&myTheme{})

	//新建一个窗口
	w := a.NewWindow("批量创建目录工具 V0.1版")

	//主界面框架布局
	MainShow(w)
	//尺寸
	w.Resize(fyne.Size{Width: 600, Height: 400})
	//w居中显示
	w.CenterOnScreen()
	//循环运行
	w.ShowAndRun()
	err := os.Unsetenv("FYNE_FONT")
	if err != nil {
		return
	}

}

/**
判断文件是否存在
*/
func IsDir(fileAddr string) bool {
	s, err := os.Stat(fileAddr)
	if err != nil {
		//fmt.Println(err)
		return false
	}
	return s.IsDir()
}

func GetCurrentDirectory() string {
	//返回绝对路径  filepath.Dir(os.Args[0])去除最后一个元素的路径
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		fmt.Println(err)
	}

	//将\替换成/
	return strings.Replace(dir, "\\", "/", -1)
}

func WalkDir(filepath string, level int) ([]string, error) {
	files, err := ioutil.ReadDir(filepath) // files为当前目录下的所有文件名称【包括文件夹】
	if err != nil {
		return nil, err
	}

	var allfile []string
	for _, v := range files {
		fullPath := filepath + "/" + v.Name() // 全路径 + 文件名称
		if v.IsDir() {                        // 如果是目录
			allfile = append(allfile, fullPath)
			a, _ := WalkDir(fullPath, level+1) // 遍历改路径下的所有文件
			allfile = append(allfile, a...)
		}
	}

	return allfile, nil
}

// scanDir 递归计算目录下所有文件
func scanDir(path string, dirMap map[string]float64) {

	dirAry, err := ioutil.ReadDir(path)
	if err != nil {
		panic(err)
	}
	for _, e := range dirAry {
		if e.IsDir() {
			scanDir(filepath.Join(path, e.Name()), dirMap)
		} else {
			dirMap[filepath.Join(path, e.Name())] = float64(e.Size()) / 1024
		}
	}
}

// MainShow 主界面函数
func MainShow(w fyne.Window) {

	Current_dir := GetCurrentDirectory()

	Excel_entry := widget.NewEntry()
	Folder_entry := widget.NewEntry()
	//设置文件夹路径为当前目录
	Folder_entry.SetText(Current_dir)

	Folder_progress := widget.NewProgressBar()
	progress_label := widget.NewLabel("进度:")

	v_progress := container.NewBorder(layout.NewSpacer(), layout.NewSpacer(), progress_label, nil, Folder_progress)
	v_progress.Hidden = true

	Excel_label := widget.NewLabel("Excel文件:")
	Folder_label := widget.NewLabel("生成目标文件夹:")

	Excel_open := widget.NewButton("选择文件", func() {
		fd := dialog.NewFileOpen(func(list fyne.URIReadCloser, err error) {
			if err != nil {
				dialog.ShowError(err, w)
				return
			}
			if list == nil {
				log.Println("取消")
				return
			}
			//out := fmt.Sprintf(list.String())
			Excel_entry.SetText(list.URI().Path())
		}, w)
		fd.SetFilter(storage.NewExtensionFileFilter([]string{".xlsx"}))
		fd.Show()
	})

	Folder_open := widget.NewButton("选择目录", func() {
		dialog.ShowFolderOpen(func(list fyne.ListableURI, err error) {
			if err != nil {
				dialog.ShowError(err, w)
				return
			}
			if list == nil {
				log.Println("取消")
				return
			}
			//out := fmt.Sprintf(list.String())
			Folder_entry.SetText(list.Path())
		}, w)
	})

	bt3 := widget.NewButton("关闭", func() {
		w.Close()
	})

	bt4 := widget.NewButton("生成目录", func() {

		if len(Excel_entry.Text) <= 0 {
			dialog.ShowInformation("提示", "请选择Excel文件!", w)
			return
		}

		if len(Folder_entry.Text) <= 0 {
			dialog.ShowInformation("提示", "请选择生成目标文件夹!", w)
			return
		}

		//创建文件目录
		file_name := Excel_entry.Text

		f, err := excelize.OpenFile(file_name)
		if err != nil {
			dialog.ShowInformation("标题", "打开Excel文件失败", w)
			return
		} else {
			defer func() {
				// Close the spreadsheet.
				if err := f.Close(); err != nil {
					fmt.Println(err)
				}
			}()

			Folder_progress.Refresh()
			v_progress.Hidden = false

			//获取Excel表格第一个Sheet
			activeSheetName := f.GetSheetList()[f.GetActiveSheetIndex()]
			separator := string(os.PathSeparator)

			rows, err := f.GetRows(activeSheetName)

			if err != nil {
				dialog.ShowInformation("标题", "打开Sheet1表格失败", w)
				return
			} else {

				Folder_progress.Min = 1
				Folder_progress.Max = float64(cap(rows))

				for r, row := range rows {

					if r >= 1 {

						_dir := Folder_entry.Text + separator

						for _, cell := range row {

							_dir += strings.TrimSpace(cell) + separator

							//fmt.Println("Sheet表名:" + activeSheetName + "===第" + strconv.Itoa(r) + "行===第" + strconv.Itoa(c) + "列===值:" + cell)
						}

						exist := IsDir(_dir)

						if exist {
							fmt.Printf("目录已存在:[%v]\n", _dir)
						} else {
							// 创建文件夹

							err := os.MkdirAll(_dir, 0755)
							if err != nil {
								fmt.Printf("创建目录失败:[%v]\n", _dir)
							} else {
								fmt.Printf("创建目录成功:[%v]\n", _dir)
							}

						}

						Folder_progress.SetValue(float64(r))

					}

				}
				Folder_progress.SetValue(Folder_progress.Max)

				dialog.ShowInformation("标题", "批量创建目录完成", w)
			}

		}

	})

	v1 := container.NewBorder(layout.NewSpacer(), layout.NewSpacer(), Excel_label, Excel_open, Excel_entry)
	v2 := container.NewBorder(layout.NewSpacer(), layout.NewSpacer(), Folder_label, Folder_open, Folder_entry)

	v3 := container.NewHBox(bt3, bt4)
	v5Center := container.NewCenter(v3)

	tab_1 := container.NewVBox(v1, v2, v_progress, v5Center)

	//==========================================================================================================

	Excel_entry_tab2 := widget.NewEntry()
	Folder_entry_tab2 := widget.NewEntry()
	//设置文件夹路径为当前目录
	Folder_entry_tab2.SetText(Current_dir)

	Folder_progress_tab2 := widget.NewProgressBar()
	progress_label_tab2 := widget.NewLabel("进度:")

	v_progress_tab2 := container.NewBorder(layout.NewSpacer(), layout.NewSpacer(), progress_label_tab2, nil, Folder_progress_tab2)
	v_progress_tab2.Hidden = true

	Excel_label_tab2 := widget.NewLabel("Excel文件:")
	Folder_label_tab2 := widget.NewLabel("目标文件夹:")

	Excel_open_tab2 := widget.NewButton("保存文件", func() {
		fd := dialog.NewFileSave(func(list fyne.URIWriteCloser, err error) {
			if err != nil {
				dialog.ShowError(err, w)
				return
			}
			if list == nil {
				log.Println("取消")
				return
			}
			//out := fmt.Sprintf(list.String())
			Excel_entry_tab2.SetText(list.URI().Path())
		}, w)
		fd.SetFilter(storage.NewExtensionFileFilter([]string{".xlsx"}))
		fd.Show()
	})

	Folder_open_tab2 := widget.NewButton("选择目录", func() {
		dialog.ShowFolderOpen(func(list fyne.ListableURI, err error) {
			if err != nil {
				dialog.ShowError(err, w)
				return
			}
			if list == nil {
				log.Println("取消")
				return
			}
			//out := fmt.Sprintf(list.String())
			Folder_entry_tab2.SetText(list.Path())
		}, w)
	})

	bt3_tab2 := widget.NewButton("关闭", func() {
		w.Close()
	})

	bt4_tab2 := widget.NewButton("生成Excel", func() {

		if len(Excel_entry_tab2.Text) <= 0 {
			dialog.ShowInformation("提示", "请选择Excel保存文件!", w)
			return
		}

		if len(Folder_entry_tab2.Text) <= 0 {
			dialog.ShowInformation("提示", "请选择目标文件夹!", w)
			return
		}

		Folder_progress_tab2.Refresh()
		v_progress_tab2.Hidden = false

		dir_name := Folder_entry_tab2.Text
		dir_list, _ := WalkDir(dir_name, 5)

		// 创建一个工作表
		f := excelize.NewFile()

		// 创建一个工作表
		index := f.NewSheet("Sheet1")

		// 设置单元格的值
		f.SetCellRichText("Sheet1", "A1", []excelize.RichTextRun{{
			Text: "第一级",
			Font: &excelize.Font{
				Bold: true,
			},
		}})
		f.SetCellRichText("Sheet1", "B1", []excelize.RichTextRun{{
			Text: "第二级",
			Font: &excelize.Font{
				Bold: true,
			},
		}})
		f.SetCellRichText("Sheet1", "C1", []excelize.RichTextRun{{
			Text: "第三级",
			Font: &excelize.Font{
				Bold: true,
			},
		}})
		f.SetCellRichText("Sheet1", "D1", []excelize.RichTextRun{{
			Text: "文件数量",
			Font: &excelize.Font{
				Bold: true,
			},
		}})

		Folder_progress_tab2.Min = 1
		Folder_progress_tab2.Max = float64(cap(dir_list))

		k := 2

		for i, v := range dir_list {

			dir_v := strings.Replace(v, dir_name, "", -1)
			countSplit := strings.Split(dir_v, "/")
			fmt.Println(dir_v)
			//fmt.Println(dir_v, countSplit, len(countSplit))
			if len(countSplit) >= 4 {

				dirMap := make(map[string]float64)
				var fileCount int   //文件数量
				var dirSize float64 //文件夹的大小
				scanDir(v, dirMap)

				f.SetCellValue("Sheet1", "A"+strconv.Itoa(k), countSplit[1])
				f.SetCellValue("Sheet1", "B"+strconv.Itoa(k), countSplit[2])
				f.SetCellValue("Sheet1", "C"+strconv.Itoa(k), countSplit[3])

				for _, v := range dirMap {
					fileCount++
					dirSize += v
				}
				f.SetCellValue("Sheet1", "D"+strconv.Itoa(k), fileCount)
				k++
			}
			Folder_progress_tab2.SetValue(float64(i))
		}

		// 设置工作簿的默认工作表
		f.SetActiveSheet(index)

		// 根据指定路径保存文件
		if err := f.SaveAs(Excel_entry_tab2.Text); err != nil {
			dialog.ShowInformation("提示", err.Error(), w)
			return
		}

		Folder_progress_tab2.SetValue(Folder_progress_tab2.Max)

		dialog.ShowInformation("提示", "生成Excel文件为:"+Excel_entry_tab2.Text, w)
	})

	v1_tab2 := container.NewBorder(layout.NewSpacer(), layout.NewSpacer(), Excel_label_tab2, Excel_open_tab2, Excel_entry_tab2)
	v2_tab2 := container.NewBorder(layout.NewSpacer(), layout.NewSpacer(), Folder_label_tab2, Folder_open_tab2, Folder_entry_tab2)

	v3_tab2 := container.NewHBox(bt3_tab2, bt4_tab2)
	v5Center_tab2 := container.NewCenter(v3_tab2)

	tab_2 := container.NewVBox(v1_tab2, v2_tab2, v_progress_tab2, v5Center_tab2)

	tabs := container.NewAppTabs(
		container.NewTabItem("Excel批量创建目录", tab_1),
		container.NewTabItem("目录生成Excel文件", tab_2),
	)

	w.SetContent(tabs)
}
