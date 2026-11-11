package merger

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	// MaxExcelRows Excel 单个工作表最大支持行数
	MaxExcelRows = 1048576
)

// DefaultMerger 默认合并器实现
type DefaultMerger struct{}

func NewDefaultMerger() *DefaultMerger {
	return &DefaultMerger{}
}

// Run 执行合并任务
func (m *DefaultMerger) Run(ctx context.Context, cfg MergeConfig, cb ProgressCallback) (*MergeResult, error) {
	startTime := time.Now()

	// 1. 规范化源目录
	if cfg.InputDir == "" {
		return nil, fmt.Errorf("源目录不能为空")
	}
	absInputDir, err := filepath.Abs(cfg.InputDir)
	if err != nil {
		return nil, fmt.Errorf("无法解析源目录路径: %w", err)
	}
	dirInfo, err := os.Stat(absInputDir)
	if err != nil || !dirInfo.IsDir() {
		return nil, fmt.Errorf("指定路径不存在或不是目录: %s", absInputDir)
	}

	// 2. 规范化输出路径
	outPath := strings.TrimSpace(cfg.OutputPath)
	if outPath == "" {
		outPath = filepath.Join(absInputDir, fmt.Sprintf("合并结果_%s.xlsx", time.Now().Format("20060102_150405")))
	}
	absOutPath, err := filepath.Abs(outPath)
	if err != nil {
		return nil, fmt.Errorf("无法解析输出文件路径: %w", err)
	}
	if strings.ToLower(filepath.Ext(absOutPath)) != ".xlsx" {
		absOutPath += ".xlsx"
	}

	// 3. 规范化参数默认值
	if cfg.HeaderRows < 0 {
		cfg.HeaderRows = 0
	}
	if len(cfg.Extensions) == 0 {
		cfg.Extensions = []string{".xlsx", ".xlsm", ".xls", ".csv"}
	}
	targetExts := make(map[string]bool)
	for _, ext := range cfg.Extensions {
		e := strings.ToLower(strings.TrimSpace(ext))
		if !strings.HasPrefix(e, ".") {
			e = "." + e
		}
		targetExts[e] = true
	}

	// 4. 扫描待处理文件
	cbLog(cb, LogLevelInfo, fmt.Sprintf("开始扫描目录: %s", absInputDir))
	matchedFiles, err := scanFiles(absInputDir, cfg.Recurse, targetExts, absOutPath)
	if err != nil {
		return nil, fmt.Errorf("扫描目录出错: %w", err)
	}

	totalFiles := len(matchedFiles)
	if totalFiles == 0 {
		cbLog(cb, LogLevelWarn, fmt.Sprintf("未在目录中找到符合条件的表格文件: %s", absInputDir))
		return &MergeResult{
			OutputPath: absOutPath,
			TotalFiles: 0,
			Duration:   time.Since(startTime),
		}, nil
	}

	cbLog(cb, LogLevelInfo, fmt.Sprintf("找到 %d 个待合并表格文件，表头行数 = %d", totalFiles, cfg.HeaderRows))

	// 5. 初始化 Excel 输出文件
	outF := excelize.NewFile()
	defer func() {
		_ = outF.Close()
	}()

	const resultSheetName = "合并结果"
	defaultSheet := outF.GetSheetName(0)
	if defaultSheet != "" {
		_ = outF.SetSheetName(defaultSheet, resultSheetName)
	} else {
		_, _ = outF.NewSheet(resultSheetName)
	}

	nextRow := 1
	headerWritten := false
	baseCols := 0
	mergedSheets := 0
	srcOffset := 0
	if cfg.AddSourceColumns {
		srcOffset = 2
	}

	// 6. 逐个文件合并处理
	for fIndex, fPath := range matchedFiles {
		// 检查上下文取消
		select {
		case <-ctx.Done():
			cbLog(cb, LogLevelWarn, "用户已取消合并操作")
			return nil, ctx.Err()
		default:
		}

		fileName := filepath.Base(fPath)
		cbProgress(cb, MergeStats{
			TotalFiles:   totalFiles,
			CurrentFile:  fileName,
			FileIndex:    fIndex + 1,
			MergedSheets: mergedSheets,
			TotalRows:    nextRow - 1,
			Duration:     time.Since(startTime),
		})

		reader, err := NewTableReader(fPath)
		if err != nil {
			cbLog(cb, LogLevelWarn, fmt.Sprintf("无法读取文件 [%s]: %v", fileName, err))
			continue
		}

		sheetNames, err := reader.GetSheets()
		if err != nil {
			_ = reader.Close()
			cbLog(cb, LogLevelWarn, fmt.Sprintf("获取工作表列表失败 [%s]: %v", fileName, err))
			continue
		}

		// 挑选目标工作表
		var targetSheets []string
		for sIdx, sName := range sheetNames {
			if isSheetSelected(sName, sIdx, &cfg) {
				targetSheets = append(targetSheets, sName)
			}
		}

		if len(targetSheets) == 0 {
			_ = reader.Close()
			cbLog(cb, LogLevelWarn, fmt.Sprintf("[%s] 未找到符合条件的工作表，已跳过。全部工作表: %s", fileName, strings.Join(sheetNames, ", ")))
			continue
		}

		// 处理每个选中的工作表
		for _, sName := range targetSheets {
			select {
			case <-ctx.Done():
				_ = reader.Close()
				return nil, ctx.Err()
			default:
			}

			rows, err := reader.ReadSheet(sName)
			if err != nil {
				cbLog(cb, LogLevelWarn, fmt.Sprintf("读取工作表失败 [%s/%s]: %v", fileName, sName, err))
				continue
			}

			totalSheetRows := len(rows)
			if totalSheetRows == 0 {
				cbLog(cb, LogLevelInfo, fmt.Sprintf("[%s/%s] 为空工作表，已跳过", fileName, sName))
				continue
			}

			totalSheetCols := len(rows[0])

			// 计算写入计划
			skipRows := 0
			if headerWritten {
				skipRows = cfg.HeaderRows
			}
			if skipRows >= totalSheetRows {
				cbLog(cb, LogLevelInfo, fmt.Sprintf("[%s/%s] 只有表头（总共 %d 行，表头 %d 行），已跳过", fileName, sName, totalSheetRows, cfg.HeaderRows))
				continue
			}

			// 本次写入的数据行数（若首表，包含表头行与后续数据行）
			rowsToWrite := totalSheetRows - skipRows
			if nextRow+rowsToWrite-1 > MaxExcelRows {
				_ = reader.Close()
				return nil, fmt.Errorf("合并后的行数超过 Excel 上限(%d 行)，已在 [%s/%s] 处停止", MaxExcelRows, fileName, sName)
			}

			// 列数一致性检测
			if !headerWritten {
				baseCols = totalSheetCols
			} else if totalSheetCols != baseCols {
				cbLog(cb, LogLevelWarn, fmt.Sprintf("[%s/%s] 列数为 %d，与首个表格的 %d 列不一致，请注意核对！", fileName, sName, totalSheetCols, baseCols))
			}

			// 写入行数据
			for rOffset := 0; rOffset < rowsToWrite; rOffset++ {
				srcRowIdx := skipRows + rOffset
				srcRow := rows[srcRowIdx]
				curDestRow := nextRow + rOffset

				// 构造写入行
				rowVals := make([]interface{}, srcOffset+len(srcRow))

				if cfg.AddSourceColumns {
					if !headerWritten && srcRowIdx < cfg.HeaderRows {
						// 首个表格的表头行
						if srcRowIdx == cfg.HeaderRows-1 {
							rowVals[0] = "来源文件"
							rowVals[1] = "来源工作表"
						} else {
							rowVals[0] = ""
							rowVals[1] = ""
						}
					} else {
						// 正常数据行
						rowVals[0] = fileName
						rowVals[1] = sName
					}
				}

				for cIdx, val := range srcRow {
					rowVals[srcOffset+cIdx] = val
				}

				cellCoord, err := excelize.CoordinatesToCellName(1, curDestRow)
				if err == nil {
					_ = outF.SetSheetRow(resultSheetName, cellCoord, &rowVals)
				}
			}

			nextRow += rowsToWrite
			headerWritten = true
			mergedSheets++
			cbLog(cb, LogLevelInfo, fmt.Sprintf("  + [%s] [%s] 成功合并 %d 行", fileName, sName, rowsToWrite))
		}

		_ = reader.Close()
	}

	if !headerWritten {
		cbLog(cb, LogLevelWarn, "未能从任何文件中提取到有效数据，未生成结果文件。")
		return &MergeResult{
			OutputPath:   absOutPath,
			TotalFiles:   totalFiles,
			MergedSheets: 0,
			TotalRows:    0,
			Duration:     time.Since(startTime),
		}, nil
	}

	// 7. 适度优化列宽
	adjustColumnWidths(outF, resultSheetName, baseCols+srcOffset)

	// 8. 保存输出文件
	cbLog(cb, LogLevelInfo, fmt.Sprintf("正在保存合并结果至: %s", absOutPath))
	if err := outF.SaveAs(absOutPath); err != nil {
		return nil, fmt.Errorf("保存结果文件失败: %w", err)
	}

	totalRows := nextRow - 1
	cbLog(cb, LogLevelSuccess, fmt.Sprintf("合并成功！共合并 %d 个工作表，总计 %d 行数据，耗时 %.2f 秒",
		mergedSheets, totalRows, time.Since(startTime).Seconds()))
	cbLog(cb, LogLevelSuccess, fmt.Sprintf("输出文件: %s", absOutPath))

	cbProgress(cb, MergeStats{
		TotalFiles:   totalFiles,
		CurrentFile:  "完成",
		FileIndex:    totalFiles,
		MergedSheets: mergedSheets,
		TotalRows:    totalRows,
		Duration:     time.Since(startTime),
	})

	return &MergeResult{
		OutputPath:   absOutPath,
		TotalFiles:   totalFiles,
		MergedSheets: mergedSheets,
		TotalRows:    totalRows,
		Duration:     time.Since(startTime),
	}, nil
}

// scanFiles 扫描目录文件
func scanFiles(dir string, recurse bool, targetExts map[string]bool, excludePath string) ([]string, error) {
	var results []string
	cleanExclude := filepath.Clean(strings.ToLower(excludePath))

	if !recurse {
		entries, err := os.ReadDir(dir)
		if err != nil {
			return nil, err
		}
		for _, e := range entries {
			if e.IsDir() {
				continue
			}
			fName := e.Name()
			if isTempOrIgnored(fName) {
				continue
			}
			ext := strings.ToLower(filepath.Ext(fName))
			if targetExts[ext] {
				p := filepath.Join(dir, fName)
				if filepath.Clean(strings.ToLower(p)) != cleanExclude {
					results = append(results, p)
				}
			}
		}
	} else {
		err := filepath.Walk(dir, func(p string, info os.FileInfo, err error) error {
			if err != nil {
				return nil
			}
			if info.IsDir() {
				return nil
			}
			fName := info.Name()
			if isTempOrIgnored(fName) {
				return nil
			}
			ext := strings.ToLower(filepath.Ext(fName))
			if targetExts[ext] {
				if filepath.Clean(strings.ToLower(p)) != cleanExclude {
					results = append(results, p)
				}
			}
			return nil
		})
		if err != nil {
			return nil, err
		}
	}

	sort.Strings(results)
	return results, nil
}

func isTempOrIgnored(name string) bool {
	return strings.HasPrefix(name, "~$") || strings.HasPrefix(name, ".~")
}

// isSheetSelected 判断工作表是否匹配规则
func isSheetSelected(sheetName string, index int, cfg *MergeConfig) bool {
	if cfg.SheetFilter == SheetFilterPattern {
		if len(cfg.SheetPatterns) == 0 {
			return true // 模式筛选但未给具体名称，则视为全选
		}
		lowerName := strings.ToLower(strings.TrimSpace(sheetName))
		for _, pattern := range cfg.SheetPatterns {
			p := strings.ToLower(strings.TrimSpace(pattern))
			if p == "" {
				continue
			}
			// 精确匹配
			if lowerName == p {
				return true
			}
			// 通配符匹配
			matched, err := filepath.Match(p, lowerName)
			if err == nil && matched {
				return true
			}
			// 支持常见子串包含
			if strings.Contains(lowerName, p) {
				return true
			}
		}
		return false
	}
	// 默认 SheetFilterFirst：仅首个工作表（index == 0）
	return index == 0
}

func adjustColumnWidths(f *excelize.File, sheet string, maxCols int) {
	for col := 1; col <= maxCols; col++ {
		colName, err := excelize.ColumnNumberToName(col)
		if err == nil {
			_ = f.SetColWidth(sheet, colName, colName, 16)
		}
	}
}

func cbLog(cb ProgressCallback, level LogLevel, msg string) {
	if cb != nil {
		cb.OnLog(level, msg)
	}
}

func cbProgress(cb ProgressCallback, stats MergeStats) {
	if cb != nil {
		cb.OnProgress(stats)
	}
}
