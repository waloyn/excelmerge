package ui

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"excelmerge/internal/merger"

	"github.com/lxn/walk"
	. "github.com/lxn/walk/declarative"
)

// AppUI 主界面控制器
type AppUI struct {
	mainWindow *walk.MainWindow

	// 路径相关控件
	inDirEdit    *walk.LineEdit
	outPathEdit  *walk.LineEdit
	browseDirBtn *walk.PushButton
	browseOutBtn *walk.PushButton

	// 合并规则控件
	headerSpin    *walk.NumberEdit
	rbFirstSheet  *walk.RadioButton
	rbAllSheets   *walk.RadioButton
	rbMatchSheets *walk.RadioButton
	sheetPatEdit  *walk.LineEdit
	cbRecurse     *walk.CheckBox
	cbAddSource   *walk.CheckBox

	// 执行与状态控件
	startBtn   *walk.PushButton
	cancelBtn  *walk.PushButton
	pBar       *walk.ProgressBar
	statusText *walk.Label
	logText    *walk.TextEdit

	// 运行时任务状态
	cancelFunc context.CancelFunc
	runningMu  sync.Mutex
	isMerging  bool
}

// NewAppUI 创建主界面控制器
func NewAppUI() *AppUI {
	return &AppUI{}
}

// Run 启动并运行界面主循环
func (ui *AppUI) Run() error {
	font := Font{Family: "Microsoft YaHei", PointSize: 9}
	titleFont := Font{Family: "Microsoft YaHei", PointSize: 9, Bold: true}

	err := MainWindow{
		AssignTo: &ui.mainWindow,
		Title:    "Excel 表格合并工具 (Win7兼容版)",
		Font:     font,
		MinSize:  Size{Width: 640, Height: 600},
		Size:     Size{Width: 700, Height: 680},
		Layout:   VBox{Margins: Margins{Left: 12, Top: 12, Right: 12, Bottom: 12}, Spacing: 10},
		Children: []Widget{
			// ======================= 卡片 1：目录与文件 =======================
			GroupBox{
				Title:  " 1. 路径设置 (支持直接拖拽文件夹到窗口) ",
				Font:   titleFont,
				Layout: Grid{Columns: 3, Spacing: 8},
				Children: []Widget{
					Label{Text: "待合并目录:", Font: font},
					LineEdit{
						AssignTo: &ui.inDirEdit,
						Font:     font,
						OnTextChanged: func() {
							ui.autoUpdateOutputPath()
						},
					},
					PushButton{
						AssignTo: &ui.browseDirBtn,
						Text:     "浏览目录...",
						Font:     font,
						OnClicked: func() {
							ui.onBrowseInputDir()
						},
					},

					Label{Text: "输出文件路径:", Font: font},
					LineEdit{
						AssignTo: &ui.outPathEdit,
						Font:     font,
					},
					PushButton{
						AssignTo: &ui.browseOutBtn,
						Text:     "更改路径...",
						Font:     font,
						OnClicked: func() {
							ui.onBrowseOutputPath()
						},
					},
				},
			},

			// ======================= 卡片 2：规则与过滤 =======================
			GroupBox{
				Title:  " 2. 合并规则与选项 ",
				Font:   titleFont,
				Layout: VBox{Spacing: 8},
				Children: []Widget{
					Composite{
						Layout: HBox{MarginsZero: true, Spacing: 12},
						Children: []Widget{
							Label{Text: "表头行数 (每表保留行数):", Font: font},
							NumberEdit{
								AssignTo: &ui.headerSpin,
								Font:     font,
								Value:    1,
								Decimals: 0,
								MinValue: 0,
								MaxValue: 1000,
							},
							HSpacer{},
							CheckBox{
								AssignTo: &ui.cbRecurse,
								Text:     "包含子目录文件",
								Font:     font,
								Checked:  false,
							},
							CheckBox{
								AssignTo: &ui.cbAddSource,
								Text:     "在首部增加[来源文件/工作表]列",
								Font:     font,
								Checked:  true,
							},
						},
					},
					Composite{
						Layout: HBox{MarginsZero: true, Spacing: 10},
						Children: []Widget{
							Label{Text: "工作表(Sheet)选择:", Font: font},
							RadioButton{
								AssignTo: &ui.rbFirstSheet,
								Text:     "仅首个工作表",
								Font:     font,
								Value:    0,
								OnClicked: func() {
									ui.sheetPatEdit.SetEnabled(false)
								},
							},
							RadioButton{
								AssignTo: &ui.rbAllSheets,
								Text:     "所有工作表",
								Font:     font,
								Value:    1,
								OnClicked: func() {
									ui.sheetPatEdit.SetEnabled(false)
								},
							},
							RadioButton{
								AssignTo: &ui.rbMatchSheets,
								Text:     "匹配名称/通配符:",
								Font:     font,
								Value:    2,
								OnClicked: func() {
									ui.sheetPatEdit.SetEnabled(true)
								},
							},
							LineEdit{
								AssignTo: &ui.sheetPatEdit,
								Font:     font,
								Enabled:  false,
							},
						},
					},
				},
			},

			// ======================= 卡片 3：操作与进度 =======================
			GroupBox{
				Title:  " 3. 执行控制与状态 ",
				Font:   titleFont,
				Layout: VBox{Spacing: 8},
				Children: []Widget{
					Composite{
						Layout: HBox{MarginsZero: true, Spacing: 10},
						Children: []Widget{
							PushButton{
								AssignTo:  &ui.startBtn,
								Text:      "开始合并表格",
								Font:      titleFont,
								MinSize:   Size{Width: 120, Height: 32},
								OnClicked: ui.onStartMerge,
							},
							PushButton{
								AssignTo:  &ui.cancelBtn,
								Text:      "终止合并",
								Font:      font,
								Enabled:   false,
								MinSize:   Size{Width: 80, Height: 32},
								OnClicked: ui.onCancelMerge,
							},
							HSpacer{},
							Label{
								AssignTo: &ui.statusText,
								Text:     "就绪。请选择待合并目录并点击开始。",
								Font:     font,
							},
						},
					},
					ProgressBar{
						AssignTo: &ui.pBar,
						MinValue: 0,
						MaxValue: 100,
						Value:    0,
					},
				},
			},

			// ======================= 卡片 4：运行日志 =======================
			GroupBox{
				Title:  " 4. 实时日志记录 ",
				Font:   titleFont,
				Layout: VBox{Spacing: 6},
				Children: []Widget{
					TextEdit{
						AssignTo: &ui.logText,
						ReadOnly: true,
						VScroll:  true,
						Font:     Font{Family: "Consolas", PointSize: 9},
					},
					Composite{
						Layout: HBox{MarginsZero: true},
						Children: []Widget{
							HSpacer{},
							PushButton{
								Text: "清空日志",
								Font: font,
								OnClicked: func() {
									ui.logText.SetText("")
								},
							},
						},
					},
				},
			},
		},
	}.Create()

	if err != nil {
		return fmt.Errorf("创建窗口失败: %w", err)
	}

	// 默认选中“仅首个工作表”
	ui.rbFirstSheet.SetChecked(true)

	// 开启文件/文件夹原生拖拽支持
	ui.mainWindow.DropFiles().Attach(func(files []string) {
		if len(files) == 0 {
			return
		}
		first := files[0]
		stat, err := os.Stat(first)
		if err == nil {
			targetDir := first
			if !stat.IsDir() {
				targetDir = filepath.Dir(first)
			}
			ui.inDirEdit.SetText(targetDir)
			ui.appendLog(merger.LogLevelInfo, fmt.Sprintf("已通过拖拽载入目录: %s", targetDir))
		}
	})

	ui.appendLog(merger.LogLevelInfo, "程序初始化完成，支持拖拽文件夹至窗口。")
	ui.mainWindow.Run()
	return nil
}

// 自动生成输出文件名
func (ui *AppUI) autoUpdateOutputPath() {
	inDir := strings.TrimSpace(ui.inDirEdit.Text())
	if inDir == "" {
		return
	}
	curOut := strings.TrimSpace(ui.outPathEdit.Text())
	if curOut == "" || strings.Contains(curOut, "合并结果_") {
		autoName := fmt.Sprintf("合并结果_%s.xlsx", time.Now().Format("20060102_150405"))
		ui.outPathEdit.SetText(filepath.Join(inDir, autoName))
	}
}

// 浏览待合并目录
func (ui *AppUI) onBrowseInputDir() {
	dlg := new(walk.FileDialog)
	dlg.Title = "选择待合并表格所在文件夹"
	if ok, err := dlg.ShowBrowseFolder(ui.mainWindow); err == nil && ok {
		ui.inDirEdit.SetText(dlg.FilePath)
	}
}

// 浏览另存为输出文件
func (ui *AppUI) onBrowseOutputPath() {
	dlg := new(walk.FileDialog)
	dlg.Title = "选择合并结果保存路径"
	dlg.Filter = "Excel 工作簿 (*.xlsx)|*.xlsx"
	if ok, err := dlg.ShowSave(ui.mainWindow); err == nil && ok {
		p := dlg.FilePath
		if !strings.HasSuffix(strings.ToLower(p), ".xlsx") {
			p += ".xlsx"
		}
		ui.outPathEdit.SetText(p)
	}
}

// 开始合并
func (ui *AppUI) onStartMerge() {
	ui.runningMu.Lock()
	if ui.isMerging {
		ui.runningMu.Unlock()
		return
	}
	ui.isMerging = true
	ui.runningMu.Unlock()

	inDir := strings.TrimSpace(ui.inDirEdit.Text())
	if inDir == "" {
		walk.MsgBox(ui.mainWindow, "提示", "请先选择待合并的文件夹路径！", walk.MsgBoxIconWarning)
		ui.setMergingState(false)
		return
	}

	outPath := strings.TrimSpace(ui.outPathEdit.Text())
	if outPath == "" {
		ui.autoUpdateOutputPath()
		outPath = strings.TrimSpace(ui.outPathEdit.Text())
	}

	headerRows := int(ui.headerSpin.Value())
	recurse := ui.cbRecurse.Checked()
	addSource := ui.cbAddSource.Checked()

	var filterMode merger.SheetFilterMode = merger.SheetFilterFirst
	var sheetPatterns []string

	if ui.rbAllSheets.Checked() {
		filterMode = merger.SheetFilterPattern
		sheetPatterns = nil // 空模式即代表全选
	} else if ui.rbMatchSheets.Checked() {
		filterMode = merger.SheetFilterPattern
		raw := strings.TrimSpace(ui.sheetPatEdit.Text())
		if raw != "" {
			parts := strings.FieldsFunc(raw, func(r rune) bool {
				return r == ',' || r == ';' || r == '，' || r == '；'
			})
			for _, p := range parts {
				sheetPatterns = append(sheetPatterns, strings.TrimSpace(p))
			}
		}
	}

	cfg := merger.MergeConfig{
		InputDir:         inDir,
		OutputPath:       outPath,
		HeaderRows:       headerRows,
		SheetFilter:      filterMode,
		SheetPatterns:    sheetPatterns,
		Recurse:          recurse,
		AddSourceColumns: addSource,
		Extensions:       []string{".xlsx", ".xlsm", ".xls", ".csv"},
	}

	ui.setMergingState(true)
	ctx, cancel := context.WithCancel(context.Background())
	ui.cancelFunc = cancel

	go func() {
		defer ui.setMergingState(false)

		m := merger.NewDefaultMerger()
		res, err := m.Run(ctx, cfg, ui)

		ui.mainWindow.Synchronize(func() {
			if err != nil {
				if err == context.Canceled {
					ui.statusText.SetText("操作已取消。")
					walk.MsgBox(ui.mainWindow, "提示", "合并操作已终止。", walk.MsgBoxIconInformation)
				} else {
					ui.statusText.SetText("合并失败！")
					ui.appendLog(merger.LogLevelError, fmt.Sprintf("执行失败: %v", err))
					walk.MsgBox(ui.mainWindow, "合并出错", err.Error(), walk.MsgBoxIconError)
				}
			} else {
				ui.pBar.SetValue(100)
				ui.statusText.SetText(fmt.Sprintf("合并完成: 共 %d 个工作表，%d 行数据。", res.MergedSheets, res.TotalRows))
				msg := fmt.Sprintf("表格合并完成！\n\n共处理文件: %d 个\n合并工作表: %d 个\n汇总数据行: %d 行\n耗时: %.1f 秒\n\n结果已保存至:\n%s",
					res.TotalFiles, res.MergedSheets, res.TotalRows, res.Duration.Seconds(), res.OutputPath)
				walk.MsgBox(ui.mainWindow, "完成", msg, walk.MsgBoxIconInformation)
			}
		})
	}()
}

// 终止合并
func (ui *AppUI) onCancelMerge() {
	if ui.cancelFunc != nil {
		ui.cancelFunc()
		ui.appendLog(merger.LogLevelWarn, "正在向任务发送终止指令...")
	}
}

// 界面控件状态切换
func (ui *AppUI) setMergingState(merging bool) {
	ui.mainWindow.Synchronize(func() {
		ui.runningMu.Lock()
		ui.isMerging = merging
		ui.runningMu.Unlock()

		ui.startBtn.SetEnabled(!merging)
		ui.cancelBtn.SetEnabled(merging)
		ui.browseDirBtn.SetEnabled(!merging)
		ui.browseOutBtn.SetEnabled(!merging)
		ui.inDirEdit.SetReadOnly(merging)
		ui.outPathEdit.SetReadOnly(merging)

		if merging {
			ui.pBar.SetValue(0)
			ui.statusText.SetText("正在处理中，请稍候...")
		}
	})
}

// OnLog 进度回调
func (ui *AppUI) OnLog(level merger.LogLevel, msg string) {
	ui.mainWindow.Synchronize(func() {
		ui.appendLog(level, msg)
	})
}

// OnProgress 进度与统计回调
func (ui *AppUI) OnProgress(stats merger.MergeStats) {
	ui.mainWindow.Synchronize(func() {
		if stats.TotalFiles > 0 {
			percent := int(float64(stats.FileIndex) / float64(stats.TotalFiles) * 100)
			if percent > 100 {
				percent = 100
			}
			ui.pBar.SetValue(percent)
			ui.statusText.SetText(fmt.Sprintf("[%d/%d] 正在处理: %s (已合并 %d 行)",
				stats.FileIndex, stats.TotalFiles, stats.CurrentFile, stats.TotalRows))
		}
	})
}

func (ui *AppUI) appendLog(level merger.LogLevel, msg string) {
	tStr := time.Now().Format("15:04:05")
	tag := fmt.Sprintf("[%s][%s]", tStr, level.String())
	line := fmt.Sprintf("%-16s %s\r\n", tag, msg)
	ui.logText.AppendText(line)
}
