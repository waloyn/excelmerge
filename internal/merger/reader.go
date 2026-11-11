package merger

import (
	"bytes"
	"encoding/csv"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"unicode/utf8"

	"github.com/extrame/xls"
	"github.com/xuri/excelize/v2"
	"golang.org/x/text/encoding/simplifiedchinese"
	"golang.org/x/text/transform"
)

// TableReader 表格读取器统一抽象
type TableReader interface {
	GetSheets() ([]string, error)
	ReadSheet(sheetName string) ([][]string, error)
	Close() error
}

// NewTableReader 根据文件后缀创建对应格式读取器
func NewTableReader(filePath string) (TableReader, error) {
	ext := strings.ToLower(filepath.Ext(filePath))
	switch ext {
	case ".xlsx", ".xlsm":
		return newExcelizeReader(filePath)
	case ".xls":
		return newXlsReader(filePath)
	case ".csv":
		return newCsvReader(filePath)
	default:
		return nil, fmt.Errorf("不支持的文件格式: %s", ext)
	}
}

// ==========================================
// 1. ExcelizeReader (.xlsx / .xlsm)
// ==========================================

type excelizeReader struct {
	f    *excelize.File
	path string
}

func newExcelizeReader(filePath string) (*excelizeReader, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("打开 Excel 文件失败 [%s]: %w", filepath.Base(filePath), err)
	}
	return &excelizeReader{f: f, path: filePath}, nil
}

func (r *excelizeReader) GetSheets() ([]string, error) {
	return r.f.GetSheetList(), nil
}

func (r *excelizeReader) ReadSheet(sheetName string) ([][]string, error) {
	// GetRows 获取所有行，并且会自动将数字、日期格式化为可见字符串
	rows, err := r.f.GetRows(sheetName)
	if err != nil {
		return nil, err
	}
	return cleanRows(rows), nil
}

func (r *excelizeReader) Close() error {
	if r.f != nil {
		return r.f.Close()
	}
	return nil
}

// ==========================================
// 2. XlsReader (.xls 97-2003 BIFF8)
// ==========================================

type xlsReader struct {
	f    *xls.WorkBook
	path string
}

func newXlsReader(filePath string) (*xlsReader, error) {
	// 使用 utf-8 尝试解析
	xlFile, err := xls.Open(filePath, "utf-8")
	if err != nil {
		return nil, fmt.Errorf("打开 XLS 文件失败 [%s]: %w", filepath.Base(filePath), err)
	}
	return &xlsReader{f: xlFile, path: filePath}, nil
}

func (r *xlsReader) GetSheets() ([]string, error) {
	var names []string
	count := r.f.NumSheets()
	for i := 0; i < count; i++ {
		s := r.f.GetSheet(i)
		if s != nil {
			names = append(names, s.Name)
		}
	}
	return names, nil
}

func (r *xlsReader) ReadSheet(sheetName string) ([][]string, error) {
	count := r.f.NumSheets()
	var targetSheet *xls.WorkSheet
	for i := 0; i < count; i++ {
		s := r.f.GetSheet(i)
		if s != nil && s.Name == sheetName {
			targetSheet = s
			break
		}
	}
	if targetSheet == nil {
		return nil, fmt.Errorf("未找到工作表: %s", sheetName)
	}

	maxRow := int(targetSheet.MaxRow)
	var rows [][]string
	for rIdx := 0; rIdx <= maxRow; rIdx++ {
		row := targetSheet.Row(rIdx)
		if row == nil {
			rows = append(rows, []string{})
			continue
		}
		var cols []string
		lastCol := row.LastCol()
		for cIdx := 0; cIdx < lastCol; cIdx++ {
			cols = append(cols, row.Col(cIdx))
		}
		rows = append(rows, cols)
	}

	return cleanRows(rows), nil
}

func (r *xlsReader) Close() error {
	// extrame/xls 不需要显式释放文件句柄
	return nil
}

// ==========================================
// 3. CsvReader (.csv)
// ==========================================

type csvReader struct {
	path string
	rows [][]string
}

func newCsvReader(filePath string) (*csvReader, error) {
	data, err := os.ReadFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("读取 CSV 失败 [%s]: %w", filepath.Base(filePath), err)
	}

	// 智能编码识别与转码：若非 UTF-8，则回退使用 GBK 解码
	var reader io.Reader
	if utf8.Valid(data) {
		reader = bytes.NewReader(data)
	} else {
		reader = transform.NewReader(bytes.NewReader(data), simplifiedchinese.GBK.NewDecoder())
	}

	csvR := csv.NewReader(reader)
	csvR.FieldsPerRecord = -1 // 允许每行字段数不同
	csvR.LazyQuotes = true

	records, err := csvR.ReadAll()
	if err != nil {
		return nil, fmt.Errorf("解析 CSV 失败 [%s]: %w", filepath.Base(filePath), err)
	}

	return &csvReader{
		path: filePath,
		rows: cleanRows(records),
	}, nil
}

func (r *csvReader) GetSheets() ([]string, error) {
	// CSV 默认视为仅包含一个名为文件名（不含扩展名）的工作表
	base := filepath.Base(r.path)
	ext := filepath.Ext(base)
	name := strings.TrimSuffix(base, ext)
	if name == "" {
		name = "Sheet1"
	}
	return []string{name}, nil
}

func (r *csvReader) ReadSheet(sheetName string) ([][]string, error) {
	return r.rows, nil
}

func (r *csvReader) Close() error {
	return nil
}

// ==========================================
// 辅助工具：修剪尾部空行与多余空单元格
// ==========================================

func cleanRows(rows [][]string) [][]string {
	// 1. 去除末尾的空白行
	lastNonEmptyRow := -1
	for r := len(rows) - 1; r >= 0; r-- {
		row := rows[r]
		hasContent := false
		for _, cell := range row {
			if strings.TrimSpace(cell) != "" {
				hasContent = true
				break
			}
		}
		if hasContent {
			lastNonEmptyRow = r
			break
		}
	}

	if lastNonEmptyRow == -1 {
		return nil // 全空
	}

	trimmed := rows[:lastNonEmptyRow+1]

	// 2. 统计最大列宽
	maxCols := 0
	for _, row := range trimmed {
		if len(row) > maxCols {
			maxCols = len(row)
		}
	}

	// 3. 补齐每行列数，确保对齐
	result := make([][]string, len(trimmed))
	for r, row := range trimmed {
		newRow := make([]string, maxCols)
		copy(newRow, row)
		result[r] = newRow
	}

	return result
}
