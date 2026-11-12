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

// FileFormatType 文件格式类型
type FileFormatType int

const (
	FormatUnknown FileFormatType = iota
	FormatXLSX
	FormatXLS
	FormatCSV
)

// detectFileFormat 通过文件头魔数与内容特征嗅探真实文件类型
func detectFileFormat(filePath string) (FileFormatType, error) {
	f, err := os.Open(filePath)
	if err != nil {
		return FormatUnknown, fmt.Errorf("打开文件失败 [%s]: %w", filepath.Base(filePath), err)
	}
	defer f.Close()

	buf := make([]byte, 512)
	n, err := f.Read(buf)
	if err != nil && err != io.EOF {
		return FormatUnknown, fmt.Errorf("读取文件头失败 [%s]: %w", filepath.Base(filePath), err)
	}
	header := buf[:n]

	// 1. 检查 XLSX / XLSM (ZIP 压缩包格式)
	// 标准 ZIP 头部为 PK\x03\x04 (0x50 0x4B 0x03 0x04) 或空包 PK\x05\x06 / PK\x07\x08
	if len(header) >= 4 {
		if header[0] == 0x50 && header[1] == 0x4B &&
			((header[2] == 0x03 && header[3] == 0x04) ||
				(header[2] == 0x05 && header[3] == 0x06) ||
				(header[2] == 0x07 && header[3] == 0x08)) {
			return FormatXLSX, nil
		}
	}

	// 2. 检查 XLS (Excel 97-2003 BIFF8 / OLE2 复合二进制文档)
	// 魔数固定为 D0 CF 11 E0 A1 B1 1A E1
	if len(header) >= 8 {
		oleHeader := []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}
		if bytes.Equal(header[:8], oleHeader) {
			return FormatXLS, nil
		}
	}

	// 3. 检查常见文本 BOM (CSV/TSV/TXT)
	// UTF-8 BOM: EF BB BF
	if len(header) >= 3 && header[0] == 0xEF && header[1] == 0xBB && header[2] == 0xBF {
		return FormatCSV, nil
	}
	// UTF-16 LE: FF FE / UTF-16 BE: FE FF
	if len(header) >= 2 && ((header[0] == 0xFF && header[1] == 0xFE) || (header[0] == 0xFE && header[1] == 0xFF)) {
		return FormatCSV, nil
	}

	// 4. 检查是否为文本表格 (无空控制字符 0x00，且以普通文本/制表符/逗号构成)
	if isTextContent(header) {
		return FormatCSV, nil
	}

	// 5. 兜底策略：按扩展名识别
	ext := strings.ToLower(filepath.Ext(filePath))
	switch ext {
	case ".xlsx", ".xlsm":
		return FormatXLSX, nil
	case ".xls":
		return FormatXLS, nil
	case ".csv":
		return FormatCSV, nil
	}

	return FormatUnknown, nil
}

// isTextContent 粗略判断数据段是否符合文本特征（不含 NUL 二进制控制符）
func isTextContent(data []byte) bool {
	if len(data) == 0 {
		return true
	}
	for _, b := range data {
		if b == 0x00 {
			return false
		}
	}
	return true
}

// detectDelimiter 嗅探 CSV 样本中的主要分隔符（逗号、制表符、分号）
func detectDelimiter(data []byte) rune {
	sample := data
	if len(sample) > 4096 {
		sample = sample[:4096]
	}

	tabs := 0
	commas := 0
	semicolons := 0

	for _, b := range sample {
		switch b {
		case '\t':
			tabs++
		case ',':
			commas++
		case ';':
			semicolons++
		}
	}

	if tabs > commas && tabs > semicolons {
		return '\t'
	}
	if semicolons > commas && semicolons > tabs {
		return ';'
	}
	return ','
}

// NewTableReader 根据文件内容类型与扩展名智能创建对应格式读取器
func NewTableReader(filePath string) (TableReader, error) {
	format, err := detectFileFormat(filePath)
	if err != nil {
		return nil, err
	}

	switch format {
	case FormatXLSX:
		r, err := newExcelizeReader(filePath)
		if err == nil {
			return r, nil
		}
		// 二级容错：尝试 CSV
		if csvR, csvErr := newCsvReader(filePath); csvErr == nil {
			return csvR, nil
		}
		return nil, err

	case FormatXLS:
		r, err := newXlsReader(filePath)
		if err == nil {
			return r, nil
		}
		// 二级容错：若非真正 OLE2 XLS（如老系统伪扩展名），降级尝试 CSV
		if csvR, csvErr := newCsvReader(filePath); csvErr == nil {
			return csvR, nil
		}
		return nil, err

	case FormatCSV:
		r, err := newCsvReader(filePath)
		if err == nil {
			return r, nil
		}
		// 二级容错：尝试 XLS
		if xlsR, xlsErr := newXlsReader(filePath); xlsErr == nil {
			return xlsR, nil
		}
		return nil, err

	default:
		ext := strings.ToLower(filepath.Ext(filePath))
		return nil, fmt.Errorf("不支持的文件格式: %s (扩展名: %s)", filepath.Base(filePath), ext)
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

	delimiter := detectDelimiter(data)

	csvR := csv.NewReader(reader)
	csvR.Comma = delimiter
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
