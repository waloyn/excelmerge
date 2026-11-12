package merger

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

type dummyCallback struct {
	logs []string
}

func (d *dummyCallback) OnLog(level LogLevel, msg string) {
	d.logs = append(d.logs, fmt.Sprintf("[%s] %s", level, msg))
}

func (d *dummyCallback) OnProgress(stats MergeStats) {}

func TestMerger(t *testing.T) {
	tempDir, err := os.MkdirTemp("", "merger_test_*")
	if err != nil {
		t.Fatalf("创建临时目录失败: %v", err)
	}
	defer os.RemoveAll(tempDir)

	// 创建测试表格 1: file1.xlsx
	f1Path := filepath.Join(tempDir, "file1.xlsx")
	f1 := excelize.NewFile()
	_ = f1.SetSheetRow("Sheet1", "A1", &[]interface{}{"姓名", "年龄", "部门"})
	_ = f1.SetSheetRow("Sheet1", "A2", &[]interface{}{"张三", 25, "研发"})
	_ = f1.SetSheetRow("Sheet1", "A3", &[]interface{}{"李四", 30, "运营"})
	if err := f1.SaveAs(f1Path); err != nil {
		t.Fatalf("保存 file1.xlsx 失败: %v", err)
	}
	_ = f1.Close()

	// 创建测试表格 2: file2.xlsx
	f2Path := filepath.Join(tempDir, "file2.xlsx")
	f2 := excelize.NewFile()
	_ = f2.SetSheetRow("Sheet1", "A1", &[]interface{}{"姓名", "年龄", "部门"})
	_ = f2.SetSheetRow("Sheet1", "A2", &[]interface{}{"王五", 28, "财务"})
	if err := f2.SaveAs(f2Path); err != nil {
		t.Fatalf("保存 file2.xlsx 失败: %v", err)
	}
	_ = f2.Close()

	// 执行合并
	outPath := filepath.Join(tempDir, "result.xlsx")
	m := NewDefaultMerger()
	cfg := MergeConfig{
		InputDir:         tempDir,
		OutputPath:       outPath,
		HeaderRows:       1,
		SheetFilter:      SheetFilterFirst,
		AddSourceColumns: true,
	}

	cb := &dummyCallback{}
	res, err := m.Run(context.Background(), cfg, cb)
	if err != nil {
		t.Fatalf("合并执行失败: %v", err)
	}

	if res.MergedSheets != 2 {
		t.Errorf("预期合并 2 个工作表，实际: %d", res.MergedSheets)
	}

	// 验证合并输出内容
	outF, err := excelize.OpenFile(outPath)
	if err != nil {
		t.Fatalf("打开输出文件失败: %v", err)
	}
	defer outF.Close()

	rows, err := outF.GetRows("合并结果")
	if err != nil {
		t.Fatalf("读取合并结果工作表失败: %v", err)
	}

	// 预期行数：1 行表头 + file1 的 2 行 + file2 的 1 行 = 4 行
	if len(rows) != 4 {
		t.Fatalf("预期 4 行数据，实际读取 %d 行", len(rows))
	}

	// 校验表头：来源文件、来源工作表、姓名、年龄、部门
	header := rows[0]
	if len(header) < 5 {
		t.Fatalf("表头列数不足: %v", header)
	}
	if header[0] != "来源文件" || header[1] != "来源工作表" || header[2] != "姓名" {
		t.Errorf("表头字段不符合预期: %v", header)
	}

	// 校验数据行 1 (张三)
	r1 := rows[1]
	if r1[0] != "file1.xlsx" || r1[1] != "Sheet1" || r1[2] != "张三" {
		t.Errorf("第 1 行数据不符合预期: %v", r1)
	}

	// 校验数据行 3 (王五)
	r3 := rows[3]
	if r3[0] != "file2.xlsx" || r3[1] != "Sheet1" || r3[2] != "王五" {
		t.Errorf("第 3 行数据不符合预期: %v", r3)
	}
}

func TestMerger_Recurse_CSV_And_SheetFilter(t *testing.T) {
	tempDir, err := os.MkdirTemp("", "merger_test2_*")
	if err != nil {
		t.Fatalf("创建临时目录失败: %v", err)
	}
	defer os.RemoveAll(tempDir)

	subDir := filepath.Join(tempDir, "sub")
	if err := os.Mkdir(subDir, 0755); err != nil {
		t.Fatalf("创建子目录失败: %v", err)
	}

	// 在根目录创建 CSV 文件
	csvPath := filepath.Join(tempDir, "data.csv")
	csvContent := "商品,单价,数量\n苹果,5,10\n香蕉,3,20\n"
	if err := os.WriteFile(csvPath, []byte(csvContent), 0644); err != nil {
		t.Fatalf("写入 CSV 失败: %v", err)
	}

	// 在子目录创建带有特定 Sheet 名称的 xlsx
	subXlsxPath := filepath.Join(subDir, "order.xlsx")
	subF := excelize.NewFile()
	targetSheet := "销售明细"
	_, _ = subF.NewSheet(targetSheet)
	_ = subF.SetSheetRow(targetSheet, "A1", &[]interface{}{"商品", "单价", "数量"})
	_ = subF.SetSheetRow(targetSheet, "A2", &[]interface{}{"橙子", 8, 15})
	if err := subF.SaveAs(subXlsxPath); err != nil {
		t.Fatalf("保存 order.xlsx 失败: %v", err)
	}
	_ = subF.Close()

	// 运行合并（开启递归与通配符筛选）
	outPath := filepath.Join(tempDir, "merged_output.xlsx")
	m := NewDefaultMerger()
	cfg := MergeConfig{
		InputDir:         tempDir,
		OutputPath:       outPath,
		HeaderRows:       1,
		SheetFilter:      SheetFilterPattern,
		SheetPatterns:    []string{"销售*", "data"},
		Recurse:          true,
		AddSourceColumns: false,
	}

	cb := &dummyCallback{}
	res, err := m.Run(context.Background(), cfg, cb)
	if err != nil {
		t.Fatalf("递归与多格式合并执行失败: %v", err)
	}

	if res.MergedSheets != 2 {
		t.Fatalf("预期合并 2 个工作表，实际: %d", res.MergedSheets)
	}

	outF, err := excelize.OpenFile(outPath)
	if err != nil {
		t.Fatalf("打开输出失败: %v", err)
	}
	defer outF.Close()

	rows, err := outF.GetRows("合并结果")
	if err != nil {
		t.Fatalf("读取合并结果失败: %v", err)
	}

	// 预期行数：1 行表头 + CSV 2 行 + 订单 1 行 = 4 行
	if len(rows) != 4 {
		t.Fatalf("预期 4 行数据，实际: %d 行", len(rows))
	}
}

func TestExampleXlsDetectionAndRead(t *testing.T) {
	examplePath := filepath.Join("..", "..", "example.xls")
	if _, err := os.Stat(examplePath); err != nil {
		t.Skip("example.xls 不存在，跳过测试")
	}

	// 1. 验证格式嗅探
	format, err := detectFileFormat(examplePath)
	if err != nil {
		t.Fatalf("detectFileFormat 报错: %v", err)
	}
	if format != FormatCSV {
		t.Errorf("预期 example.xls 被识别为 FormatCSV (%d)，实际为: %d", FormatCSV, format)
	}

	// 2. 验证 TableReader 创建与读取
	reader, err := NewTableReader(examplePath)
	if err != nil {
		t.Fatalf("创建 TableReader 失败: %v", err)
	}
	defer reader.Close()

	sheets, err := reader.GetSheets()
	if err != nil {
		t.Fatalf("GetSheets 失败: %v", err)
	}
	if len(sheets) == 0 {
		t.Fatalf("工作表列表为空")
	}

	rows, err := reader.ReadSheet(sheets[0])
	if err != nil {
		t.Fatalf("ReadSheet 失败: %v", err)
	}

	if len(rows) < 10 {
		t.Fatalf("预期读取行数 >= 10，实际读取到 %d 行", len(rows))
	}

	// 3. 校验列数与中文内容（制表符切分与 GBK 转码校验）
	header := rows[0]
	if len(header) < 20 {
		t.Fatalf("预期字段数 >= 20，实际只有 %d 列，可能制表符未正确切分: %v", len(header), header)
	}

	if header[0] != "ShopToutID" || header[4] != "Goods_No" {
		t.Errorf("表头字段不符合预期: %v", header)
	}

	// 校验数据行包含中文正确解码（例如第 1 行包含 "连衣裙" 或 "金台绣"）
	foundChinese := false
	for _, row := range rows {
		for _, cell := range row {
			if strings.Contains(cell, "连衣裙") || strings.Contains(cell, "金台绣") || strings.Contains(cell, "女上衣") {
				foundChinese = true
				break
			}
		}
		if foundChinese {
			break
		}
	}
	if !foundChinese {
		t.Errorf("未在解析结果中找到正确解码的中文字符（可能 GBK 转码失败）")
	}
}

func TestMergerWithDisguisedXls(t *testing.T) {
	examplePath := filepath.Join("..", "..", "example.xls")
	if _, err := os.Stat(examplePath); err != nil {
		t.Skip("example.xls 不存在，跳过测试")
	}

	tempDir, err := os.MkdirTemp("", "merger_disguised_test_*")
	if err != nil {
		t.Fatalf("创建临时目录失败: %v", err)
	}
	defer os.RemoveAll(tempDir)

	// 复制一份 example.xls 到临时目录
	data, err := os.ReadFile(examplePath)
	if err != nil {
		t.Fatalf("读取 example.xls 失败: %v", err)
	}
	destXls := filepath.Join(tempDir, "test_disguised.xls")
	if err := os.WriteFile(destXls, data, 0644); err != nil {
		t.Fatalf("写入测试伪装文件失败: %v", err)
	}

	outPath := filepath.Join(tempDir, "output.xlsx")
	m := NewDefaultMerger()
	cfg := MergeConfig{
		InputDir:         tempDir,
		OutputPath:       outPath,
		HeaderRows:       1,
		SheetFilter:      SheetFilterFirst,
		AddSourceColumns: true,
	}

	cb := &dummyCallback{}
	res, err := m.Run(context.Background(), cfg, cb)
	if err != nil {
		t.Fatalf("合并执行失败: %v", err)
	}

	if res.MergedSheets != 1 {
		t.Errorf("预期合并 1 个工作表，实际: %d", res.MergedSheets)
	}

	outF, err := excelize.OpenFile(outPath)
	if err != nil {
		t.Fatalf("打开输出失败: %v", err)
	}
	defer outF.Close()

	rows, err := outF.GetRows("合并结果")
	if err != nil {
		t.Fatalf("读取合并结果失败: %v", err)
	}

	if len(rows) < 10 {
		t.Fatalf("预期合并输出行数 >= 10，实际: %d", len(rows))
	}

	// 来源列检查
	header := rows[0]
	if header[0] != "来源文件" || header[1] != "来源工作表" || header[2] != "ShopToutID" {
		t.Errorf("合并表头异常: %v", header)
	}
}

func TestSanitizeSheetName(t *testing.T) {
	used := make(map[string]bool)

	// 1. 过滤非法字符
	s1 := SanitizeSheetName("2026/11/12:报告[最新]*?.xlsx", used)
	if s1 != "2026_11_12_报告_最新___" {
		t.Errorf("过滤非法字符失败，实际: %s", s1)
	}

	// 2. 超长截断 (最大 31 个字符)
	longName := "这是一段非常非常非常非常非常非常非常长的工作表名称用于测试截断限制.xlsx"
	s2 := SanitizeSheetName(longName, used)
	if len([]rune(s2)) > 31 {
		t.Errorf("超长名称未被截断在31字内，实际字数: %d (%s)", len([]rune(s2)), s2)
	}

	// 3. 重名防冲突递增
	s3 := SanitizeSheetName("销售报表.xlsx", used)
	if s3 != "销售报表" {
		t.Errorf("首个名称不符合预期: %s", s3)
	}
	s4 := SanitizeSheetName("销售报表.csv", used)
	if s4 != "销售报表 (2)" {
		t.Errorf("重名递增不符合预期: %s", s4)
	}
	s5 := SanitizeSheetName("销售报表.xls", used)
	if s5 != "销售报表 (3)" {
		t.Errorf("第三次重名递增不符合预期: %s", s5)
	}

	// 4. 空名处理
	s6 := SanitizeSheetName("   .xlsx", used)
	if s6 != "Sheet" {
		t.Errorf("空名处理不符合预期: %s", s6)
	}
}

func TestMerger_PerFileSheet(t *testing.T) {
	tempDir, err := os.MkdirTemp("", "merger_per_file_test_*")
	if err != nil {
		t.Fatalf("创建临时目录失败: %v", err)
	}
	defer os.RemoveAll(tempDir)

	// 1. 创建 file1.xlsx，内含 2 个工作表
	f1Path := filepath.Join(tempDir, "file1.xlsx")
	f1 := excelize.NewFile()
	_ = f1.SetSheetRow("Sheet1", "A1", &[]interface{}{"姓名", "部门"})
	_ = f1.SetSheetRow("Sheet1", "A2", &[]interface{}{"张三", "研发"})
	_, _ = f1.NewSheet("Sheet2")
	_ = f1.SetSheetRow("Sheet2", "A1", &[]interface{}{"姓名", "部门"})
	_ = f1.SetSheetRow("Sheet2", "A2", &[]interface{}{"李四", "测试"})
	if err := f1.SaveAs(f1Path); err != nil {
		t.Fatalf("保存 file1.xlsx 失败: %v", err)
	}
	_ = f1.Close()

	// 2. 创建 file2.csv，单表
	f2Path := filepath.Join(tempDir, "file2.csv")
	csvContent := "姓名,部门\n王五,运营\n赵六,市场\n"
	if err := os.WriteFile(f2Path, []byte(csvContent), 0644); err != nil {
		t.Fatalf("写入 file2.csv 失败: %v", err)
	}

	// 3. 在子目录创建同名 file1.xlsx
	subDir := filepath.Join(tempDir, "sub")
	if err := os.MkdirAll(subDir, 0755); err != nil {
		t.Fatalf("创建子目录失败: %v", err)
	}
	f3Path := filepath.Join(subDir, "file1.xlsx")
	f3 := excelize.NewFile()
	_ = f3.SetSheetRow("Sheet1", "A1", &[]interface{}{"姓名", "部门"})
	_ = f3.SetSheetRow("Sheet1", "A2", &[]interface{}{"孙七", "行政"})
	if err := f3.SaveAs(f3Path); err != nil {
		t.Fatalf("保存 sub/file1.xlsx 失败: %v", err)
	}
	_ = f3.Close()

	// 执行多表输出合并 (所有工作表模式，递归子目录)
	outPath := filepath.Join(tempDir, "result_per_file.xlsx")
	m := NewDefaultMerger()
	cfg := MergeConfig{
		InputDir:         tempDir,
		OutputPath:       outPath,
		HeaderRows:       1,
		OutputMode:       OutputModePerFileSheet,
		SheetFilter:      SheetFilterPattern, // 全选工作表
		Recurse:          true,
		AddSourceColumns: true,
	}

	cb := &dummyCallback{}
	res, err := m.Run(context.Background(), cfg, cb)
	if err != nil {
		t.Fatalf("多表合并执行失败: %v", err)
	}

	if res.TotalFiles != 3 {
		t.Errorf("预期处理 3 个文件，实际: %d", res.TotalFiles)
	}
	// 合并的总工作表数：file1 有 2 个，file2 有 1 个，sub/file1 有 1 个，共 4 个
	if res.MergedSheets != 4 {
		t.Errorf("预期合并 4 个工作表，实际: %d", res.MergedSheets)
	}

	// 验证生成的工作簿内容
	outF, err := excelize.OpenFile(outPath)
	if err != nil {
		t.Fatalf("打开输出文件失败: %v", err)
	}
	defer outF.Close()

	sheets := outF.GetSheetList()
	// 预期生成的 3 个工作表分别对应 3 个文件：file1, file1 (2), file2 (排序可能受遍历顺序影响)
	if len(sheets) != 3 {
		t.Fatalf("预期输出 3 个工作表，实际工作表列表: %v", sheets)
	}

	sheetMap := make(map[string]bool)
	for _, s := range sheets {
		sheetMap[s] = true
	}

	if !sheetMap["file1"] || !sheetMap["file1 (2)"] || !sheetMap["file2"] {
		t.Errorf("工作表名称不符合预期: %v", sheets)
	}

	// 验证 file1 的多表合并：包含 Sheet1 和 Sheet2 的数据（1表头 + 2数据行 = 3行）
	f1Rows, err := outF.GetRows("file1")
	if err != nil {
		t.Fatalf("读取 file1 工作表失败: %v", err)
	}
	if len(f1Rows) != 3 {
		t.Fatalf("预期 file1 工作表有 3 行数据（含表头），实际读取: %d 行", len(f1Rows))
	}

	// 校验来源列
	if f1Rows[0][0] != "来源文件" || f1Rows[0][1] != "来源工作表" || f1Rows[0][2] != "姓名" {
		t.Errorf("file1 表头异常: %v", f1Rows[0])
	}
	if f1Rows[1][0] != "file1.xlsx" || f1Rows[1][1] != "Sheet1" || f1Rows[1][2] != "张三" {
		t.Errorf("file1 第一行数据异常: %v", f1Rows[1])
	}
	if f1Rows[2][0] != "file1.xlsx" || f1Rows[2][1] != "Sheet2" || f1Rows[2][2] != "李四" {
		t.Errorf("file1 第二行数据异常: %v", f1Rows[2])
	}
}

