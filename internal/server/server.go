package server

import (
	"log"
	"runtime"

	"github.com/mark3labs/mcp-go/server"
	"github.com/vKenjo/ms-excel-mcp-server/internal/tools"
)

type ExcelServer struct {
	server *server.MCPServer
}

func New(version string) *ExcelServer {
	s := &ExcelServer{}

	s.server = server.NewMCPServer(
		"excel-mcp-server",
		version,
	)

	// Add tools with error handling
	defer func() {
		if r := recover(); r != nil {
			log.Printf("Panic during tool registration: %v", r)
		}
	}()

	tools.AddExcelDescribeSheetsTool(s.server)
	tools.AddExcelReadSheetTool(s.server)
	
	// Only add Windows-specific tools on Windows
	if runtime.GOOS == "windows" {
		tools.AddExcelScreenCaptureTool(s.server)
	}
	
	tools.AddExcelWriteToSheetTool(s.server)
	tools.AddExcelCreateTableTool(s.server)
	tools.AddExcelCopySheetTool(s.server)
	tools.AddExcelAddDataValidationTool(s.server)
	tools.AddExcelAddConditionalFormattingTool(s.server)
	tools.AddExcelExecuteVBATool(s.server)
	tools.AddExcelAddVBAModuleTool(s.server)

	return s
}

func (s *ExcelServer) Start() error {
	defer func() {
		if r := recover(); r != nil {
			log.Printf("Panic during server start: %v", r)
		}
	}()

	return server.ServeStdio(s.server)
}
