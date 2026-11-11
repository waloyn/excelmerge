package merger

import (
	"context"
	"time"
)

// LogLevel 日志严重级别
type LogLevel int

const (
	LogLevelInfo LogLevel = iota
	LogLevelWarn
	LogLevelError
	LogLevelSuccess
)

func (l LogLevel) String() string {
	switch l {
	case LogLevelWarn:
		return "WARN"
	case LogLevelError:
		return "ERROR"
	case LogLevelSuccess:
		return "SUCCESS"
	default:
		return "INFO"
	}
}

// SheetFilterMode 工作表筛选模式
type SheetFilterMode int

const (
	// SheetFilterFirst 仅合并首个工作表（默认）
	SheetFilterFirst SheetFilterMode = iota
	// SheetFilterAll 合并所有工作表
	SheetFilterPattern
)

// MergeConfig 合并任务配置
type MergeConfig struct {
	InputDir         string          // 待合并目录
	OutputPath       string          // 输出路径（留空则自动生成）
	HeaderRows       int             // 表头行数，默认 1
	SheetFilter      SheetFilterMode // 工作表过滤模式
	SheetPatterns    []string        // 自定义工作表通配符列表（支持逗号分隔）
	Recurse          bool            // 是否递归子目录
	AddSourceColumns bool            // 是否在最左侧添加来源列（来源文件、来源工作表）
	Extensions       []string        // 匹配的文件后缀，如 [".xlsx", ".xlsm", ".xls", ".csv"]
}

// MergeStats 实时或最终统计信息
type MergeStats struct {
	TotalFiles   int
	CurrentFile  string
	FileIndex    int
	MergedSheets int
	TotalRows    int
	Duration     time.Duration
}

// MergeResult 合并最终结果
type MergeResult struct {
	OutputPath   string
	TotalFiles   int
	MergedSheets int
	TotalRows    int
	Duration     time.Duration
}

// ProgressCallback 进度与状态回调
type ProgressCallback interface {
	OnLog(level LogLevel, msg string)
	OnProgress(stats MergeStats)
}

// Merger 合并器接口
type Merger interface {
	Run(ctx context.Context, cfg MergeConfig, cb ProgressCallback) (*MergeResult, error)
}
