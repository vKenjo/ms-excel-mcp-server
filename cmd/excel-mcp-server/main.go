package main

import (
	"fmt"
	"os"

	"github.com/vKenjo/ms-excel-mcp-server/internal/server"
)

var (
  version = "dev"
)

func main() {
	s := server.New(version)
	err := s.Start()
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed to start the server: %v\n", err)
		os.Exit(1)
	}
}
