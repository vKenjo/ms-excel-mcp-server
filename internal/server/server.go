package server

import (
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
	tools.AddExcelDescribeSheetsTool(s.server)
	tools.AddExcelReadSheetTool(s.server)
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
	return server.ServeStdio(s.server)
}
